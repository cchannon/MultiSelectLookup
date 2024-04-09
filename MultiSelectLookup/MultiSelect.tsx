import * as React from "react";
import {
  Combobox,
  makeStyles,
  Option,
  tokens,
  useId,
  FluentProvider,
  SplitButton,
  Text,
  Button,
  Divider,
  Label,
  Field,
  ProgressBar
} from "@fluentui/react-components";
import {
  Dismiss12Regular,
  GlobeSearchRegular,
  AddRegular,
} from "@fluentui/react-icons";

export interface IMultiSelectProps {
  thisTableName: string;
  thisRecordId: string;
  items: ComponentFramework.PropertyTypes.DataSet;
  navigateToRecord: (tableName: string, entityId: string) => void;
  theme: any;
  utils: ComponentFramework.Utility;
  webApi: ComponentFramework.WebApi;
  currentUserId: string;
  clientURL: string;
  addNewCallback: () => void;
  allowAddNew: boolean;
  label: string | null;
  relationshipName: string;
  labelLocation: "above" | "left";
  labelMaxWidth: string;
}

const useStyles = makeStyles({
  root: {
    display: "grid",
    flexDirection: "row",
    gridTemplateRows: "repeat(1fr)",
    JustifyContent: "space-between",
    alignItems: "end",
    width: "100%",
    paddingTop: "8px",
    rowGap: "5px",
  },
  leftlabel: {
    display: "flex",
  },
  tagsList: {
    listStyleType: "none",
    marginBottom: tokens.spacingVerticalXXS,
    marginTop: 0,
    paddingTop: "3px",
    paddingLeft: 0,
    display: "flex",
    flexWrap: "wrap",
    gridGap: tokens.spacingHorizontalXXS,
  },
  wrapper: {
    columnGap: "15px",
    display: "flex",
    alignContent: "space-around",
  },
});

// save for future use

// const useDebounce = (value: any, delay: any) => {
//   const [debouncedValue, setDebouncedValue] = React.useState(value);

//   React.useEffect(() => {
//     const handler = setTimeout(() => {
//       setDebouncedValue(value);
//     }, delay);

//     return () => {
//       clearTimeout(handler);
//     };
//   }, [value, delay]);

//   return debouncedValue;
// };

export const MultiselectWithTags: React.FC<IMultiSelectProps> = (
  props: IMultiSelectProps
) => {
  //state for the metadata
  const [thisSetName, setThisSetName] = React.useState<string>("");
  const [targetSetName, setTargetSetName] = React.useState<string>("");
  const [targetPrimaryColumn, setTargetPrimaryColumn] =
    React.useState<string>("");
  const [targetCollectionName, setTargetCollectionName] =
    React.useState<string>("");
  const [targetDisplayName, setTargetDisplayName] = React.useState<string>("");

  //state for the combobox
  const [hasFocus, setHasFocus] = React.useState<boolean>(false);
  const [options, setOptions] = React.useState<
    ComponentFramework.WebApi.Entity[]
  >([]);
  const [selectedOptions, setSelectedOptions] = React.useState<string[]>([]);
  const [addNew, setAddNew] = React.useState<boolean>(false);
  const [progressBar, setProgressBar] = React.useState<boolean>(false);
  const [labelMax, setLabelMax] = React.useState<string>("140px");

  const comboboxRef = React.useRef<HTMLInputElement | null>(null);
  // save for future use
  //const debouncedSearchTerm = useDebounce(inputValue, 500); // 500ms delay
  const comboId = useId("Multiselect-Search");
  const selectedListId = `${comboId}-selection`;
  const styles = useStyles();

  React.useEffect(() => {
    if(targetPrimaryColumn === "") return;
    // TODO: get better filtering in here
    props.webApi
      .retrieveMultipleRecords(
        props.items.getTargetEntityType(),
        `?$select=${targetPrimaryColumn},${props.items.getTargetEntityType()}id&$filter=(statecode eq 0)`
      )
      .then(
        (response) => {
          setOptions(response.entities);
        },
        (error) => {
          console.error(error);
        }
      );
  }, [targetPrimaryColumn]);

  React.useEffect(() => {
    //get the primarycolumn, OTC, and setName from metadata query
    if (props.items) {
      props.utils
        .getEntityMetadata(props.items.getTargetEntityType())
        .then((metadata) => {
          setTargetPrimaryColumn(metadata["PrimaryNameAttribute"]);
          setTargetCollectionName(metadata["DisplayCollectionName"]);
          setTargetDisplayName(metadata["DisplayName"]);
          setTargetSetName(metadata["EntitySetName"]);
        });
    }
  }, [props.items]);

  React.useEffect(() => {
    if (props.thisTableName) {
      props.utils.getEntityMetadata(props.thisTableName).then((metadata) => {
        setThisSetName(metadata["EntitySetName"]);
      });
    }
  }, [props.thisTableName]);

  React.useEffect(() => {
    if (props.items.sortedRecordIds.length > 0 && targetPrimaryColumn !== "") {
      setSelectedOptions(
        props.items.sortedRecordIds.map((id) =>
          props.items.records[id].getFormattedValue(targetPrimaryColumn)
        )
      );
    }
  }, [props.items, targetPrimaryColumn]);

  //save for future use

  // React.useEffect(() => {
  //   if (debouncedSearchTerm) {
  //     //SEARCH STUFF
  //   }
  // }, [debouncedSearchTerm]);

  React.useEffect(() => {
    if (addNew) {
      if (comboboxRef.current) {
        comboboxRef.current.blur();
      }
      // document.getElementById("DS-Search-Combo")?.focus();
      // document.getElementById("DS-Search-Combo")?.blur();
      props.addNewCallback();
      setAddNew(false);
    }
  }, [addNew]);

  const onClickPrimary = (option: string) => {
    props.navigateToRecord(
      props.items.getTargetEntityType(),
      props.items.sortedRecordIds.find(
        (id) =>
          props.items.records[id].getFormattedValue(targetPrimaryColumn) ===
          option
      )!
    );
  };

  const onClickClose = (option: string) => {
    onSelectItems(selectedOptions.filter((opt) => opt !== option));
  };

  const onSelectItems = (items: string[]) => {
    let newOptions: string[] = [];
    let removedOptions: string[] = [];
    if (items.length > 0) {
      if (selectedOptions.length === 0) {
        newOptions = items;
        setProgressBar(true);
      } else {
        newOptions = items.filter((x) => !selectedOptions.includes(x));
        removedOptions = selectedOptions.filter((x) => !items.includes(x));
        setProgressBar(true);
      }
    } else {
      removedOptions = selectedOptions;
      setProgressBar(true);
    }

    setSelectedOptions(items);

    newOptions.map((opt) => {
      let id = options.find((option) => {
        if (option[targetPrimaryColumn] === opt) {
          return true;
        }
      })![`${props.items.getTargetEntityType()}id`];
      const payload = {
        "@odata.id": `${props.clientURL}/api/data/v9.1/${thisSetName}(${props.thisRecordId})`,
      };

      window
        .fetch(
          `${props.clientURL}/api/data/v9.1/${targetSetName}(${id})/${props.relationshipName}/$ref`,
          {
            method: "POST",
            headers: {
              "Content-Type": "application/json; charset=utf-8",
              Accept: "application/json",
              "OData-MaxVersion": "4.0",
              "OData-Version": "4.0",
            },
            body: JSON.stringify(payload),
          }
        )
        .then(() => cleanup(items));
    });

    removedOptions.map((opt) => {
      let id = props.items.sortedRecordIds.find(
        (x) =>
          props.items.records[x].getFormattedValue(targetPrimaryColumn) === opt
      );
      window
        .fetch(
          `${props.clientURL}/api/data/v9.1/${thisSetName}(${props.thisRecordId})/${props.relationshipName}(${id})/$ref`,
          {
            method: "DELETE",
            headers: {
              "Content-Type": "application/json; charset=utf-8",
              Accept: "application/json",
              "OData-MaxVersion": "4.0",
              "OData-Version": "4.0",
            },
          }
        )
        .then(() => cleanup(items));
    });
  };

  const cleanup = (items: string[] | undefined) => {
    setSelectedOptions(items ?? []);
    setProgressBar(false);
  };
  
  React.useEffect(() => {
    if(props.labelMaxWidth !== "")
    {
      setLabelMax(props.labelMaxWidth);
    }
  }, [props.labelMaxWidth]);

  return (
    <FluentProvider theme={props.theme} style={{ width: "100%" }}>
      <div
        className={
          props.labelLocation === "above" ? styles.root : styles.leftlabel
        }
      >
        <Label style={{ width: labelMax, paddingTop: "5px" }}>
          {props.label}
        </Label>
        <div className={styles.root} style={{width: "320px"}}>
          <Combobox
            ref={comboboxRef}
            multiselect={true}
            selectedOptions={selectedOptions}
            appearance="filled-lighter"
            aria-labelledby={comboId}
            placeholder={hasFocus ? "" : "---"}
            style={{ background: "#F5F5F5", width: "300px", paddingTop: "0px" }}
            expandIcon={<GlobeSearchRegular />}
            onFocus={(_e) => {
              setHasFocus(true);
            }}
            onBlur={(_e) => {
              setHasFocus(false);
            }}
            // save for future controlled combobox use (dataverse search)

            // onInput={(ev: React.ChangeEvent<HTMLInputElement>) => {
            //   setInputValue(ev.target.value);
            // }}
            onOptionSelect={(_event, data) => {
              onSelectItems(data.selectedOptions);
            }}
            onSelect={() => {
              console.log("Selected");
            }}
            onInputCapture={() => console.log("input capture")}
            type="search"
          >
            {targetCollectionName && (
              <Text
                style={{
                  margin: "5px",
                  padding: "5px",
                }}
              >
                All {targetCollectionName ?? ""}
              </Text>
            )}
            {options.length > 0 &&
              options.map((option) => (
                <Option key={option[targetPrimaryColumn]}>
                  {option[targetPrimaryColumn]}
                </Option>
              ))}
            {props.allowAddNew && (
              <>
                <Divider />
                <div className={styles.wrapper}>
                  <Button
                    icon={<AddRegular />}
                    appearance="subtle"
                    onClick={() => setAddNew(true)}
                  >
                    New {targetDisplayName ?? ""}
                  </Button>
                </div>
              </>
            )}
          </Combobox>
          {progressBar && (
            <Field validationMessage="saving..." validationState="none">
              <ProgressBar />
            </Field>
          )}
          {selectedOptions.length ? (
            // testing future consolidation of tags and search box
            // <div style={{background: "#F5F5F5", borderRadius:"5px", width: "100%", height: "auto", alignItems:"center"}} className={styles.leftlabel}>
              <ul id={selectedListId} className={styles.tagsList} >
                {/* The "Remove" span is used for naming the buttons without affecting the Combobox name */}
                <span id={`${comboId}-remove`} hidden>
                  Remove
                </span>
                {selectedOptions.map((option, i) => (
                  <li key={option}>
                    <SplitButton
                      size="small"
                      shape="circular"
                      appearance="primary"
                      menuButton={{
                        style: {
                          color: "rgb(17, 94, 163)",
                          background: "rgb(235, 243, 252)",
                        },
                        onClick: () => onClickClose(option),
                      }}
                      menuIcon={<Dismiss12Regular />}
                      primaryActionButton={{
                        style: {
                          color: "rgb(17, 94, 163)",
                          background: "rgb(235, 243, 252)",
                        },
                        onClick: () => onClickPrimary(option),
                      }}
                      id={`${comboId}-remove-${i}`}
                      aria-labelledby={`${comboId}-remove ${comboId}-remove-${i}`}
                    >
                      {option}
                    </SplitButton>
                  </li>
                ))}
              </ul>
              // testing future consolidation of tags and search box
              // <GlobeSearchRegular style={{height: "20px", width: "20px", flexShrink: 0, color: "#616161", paddingRight: "8px"}}/>
            // </div>
          ) : null}
        </div>
      </div>
    </FluentProvider>
  );
};
