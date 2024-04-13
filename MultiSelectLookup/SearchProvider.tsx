export interface Entity {
  name: string;
  filter?: any;
  searchColumns?: string[];
}

//base class for search providers
export class SearchProvider {
  tableName: string;
  webApi: ComponentFramework.WebApi;
  primaryColumn: string;
  filter: string | null;
  

  constructor(webApi: ComponentFramework.WebApi, tableName: string, primaryColumn: string, filter: string | null) {
    this.webApi = webApi;
    this.tableName = tableName;
    this.primaryColumn = primaryColumn;
    this.filter = filter;
  }

  initialResults = async () => {
    return([]) as ComponentFramework.WebApi.Entity[];
  } 

  search = async (query: string) => {
    return([]) as ComponentFramework.WebApi.Entity[];
  }
}

export class SimpleSearchProvider extends SearchProvider {
  records: ComponentFramework.WebApi.Entity[];

  initialResults = async(): Promise<ComponentFramework.WebApi.Entity[]> =>  {
    const res = await this.webApi
      .retrieveMultipleRecords(
        this.tableName,
        `?$select=${this.primaryColumn},${this.tableName}id${this.filter ? '&?filter='.concat(this.filter) : ""}`
      )
      .then(
        (response) => {
          this.records = response.entities;
          return (response.entities);
        },
        (error) => {
          console.error(error);
          return([]);
        }
      );
      return res;
  }

  search = async(query: string): Promise<ComponentFramework.WebApi.Entity[]> =>{
    let res = this.records.filter((record) => {
      return record[this.primaryColumn].includes(query);
    });
    return res;
  }
}

export class AdvancedSearchProvider extends SearchProvider {
  bestEffort: boolean;
  matchWords: string;
  clientURL: string;
  searchColumns: string[];

  constructor(webApi: ComponentFramework.WebApi, tableName: string, primaryColumn: string, filter: string | null, searchColumns: string[] = [], bestEffort: boolean, matchWords: string, clientURL: string) {
    super(webApi, tableName, primaryColumn, filter);
    this.bestEffort = bestEffort;
    this.matchWords = matchWords;
    this.clientURL = clientURL;
    this.searchColumns = searchColumns;
  }

  initialResults = async(): Promise<ComponentFramework.WebApi.Entity[]> =>  {
    let res = this.search("").then((results) => {
      return results;
    });
    return res;
  }

  search = async (query: string) => {
    let entities: Entity[] = [
      {
        name: this.tableName,
        filter: this.filter? this.filter : undefined,
        searchColumns: this.searchColumns.length > 0 ? this.searchColumns : undefined
      }
    ];

    const searchRequest = {
      search: query as string,
      entities: JSON.stringify(entities),
      options: JSON.stringify({
        besteffortsearchenabled: this.bestEffort,
        searchmode: this.matchWords
      }),
      count: true
    };

    // Make the POST request to the Dataverse Search API
    let res = await fetch(`${this.clientURL}/api/search/v2.0/query`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        'Accept': 'application/json'
      },
      body: JSON.stringify(searchRequest)
    })
      .then(response => response.json())
      .then(data => {
        let results = JSON.parse(data.response).Value.map((record: any) => {
          return record.Highlights.name[0];
        });
        return (results);
      })
      .catch(error => {
        // Handle the error
        console.error('Error:', error);
      });
    
      return res;
  }
}
