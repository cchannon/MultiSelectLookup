export interface Entity {
  name: string;
  filter?: string;
  searchColumns?: string[];
}

//base class for search providers
export class SearchProvider {
  tableName: string;
  webApi: ComponentFramework.WebApi;
  primaryColumn: string;
  filter: string | null;
  order: string | null;
  

  constructor(webApi: ComponentFramework.WebApi, tableName: string, primaryColumn: string, filter: string | null, order: string | null) {
    this.webApi = webApi;
    this.tableName = tableName;
    this.primaryColumn = primaryColumn;
    this.filter = filter;
    this.order = order;
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
        `?$select=${this.primaryColumn},${this.tableName}id${this.filter ? '&$filter='.concat(this.filter) : ""}${this.order ? '&$orderby='.concat(this.order) : ""}`
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
      return (record[this.primaryColumn] as string).toLowerCase().includes(query.toLowerCase());
    });
    return res;
  }
}

export class AdvancedSearchProvider extends SearchProvider {
  bestEffort: boolean;
  matchWords: string;
  clientURL: string;
  searchColumns: string[];

  constructor(webApi: ComponentFramework.WebApi, tableName: string, primaryColumn: string, filter: string | null, order: string | null, searchColumns: string[] = [], bestEffort: boolean, matchWords: string, clientURL: string) {
    super(webApi, tableName, primaryColumn, filter, order);
    this.bestEffort = bestEffort;
    this.matchWords = matchWords;
    this.clientURL = clientURL;
    this.searchColumns = searchColumns;
  }

  initialResults = async(): Promise<ComponentFramework.WebApi.Entity[]> =>  {
    let simpleProvider = new SimpleSearchProvider(this.webApi, this.tableName, this.primaryColumn, this.filter, this.order);
    let res = simpleProvider.initialResults().then((results) => {
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
      orderby: this.order? JSON.stringify([this.order]) : undefined,
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
          return record.Attributes;
        });
        return (results);
      })
      .catch(error => {
        // Handle the error
        console.error('Error:', error);
      });
    
      return res as ComponentFramework.WebApi.Entity[];
  }
}
