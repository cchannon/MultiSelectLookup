# Multi-Select Lookup

This repo contains the working code for replacing a N:N subgrid with a convincing Model Driven App Multi-Select Lookup experience. It relies heavily on Fluent UI React V9 to create a convincing "new look and feel" user experience, and lists selected tags in a grouping beneath the lookup control.

![gif demonstration of the code in action](/images/demo.gif)

## Usage
Build the control and import to your solution, or optionally just import the [solution.zip](/Solution/Solution.zip) file to your environment to get started right away.

The control provides for several configuration parameters. Some of them may be a bit surprising, but remember: we are replacing a _subgrid_ to make it look like a _field_ so some seemingly obvious elements need to be configured.

**PROPERTIES**
- dataset: this is the N:N relationship to which you will bind the control and from which it will populate 'lookup' relationships.
- Relationship Name: this is the schemaname of the N:N relationship (_not_ the intersect entity). It is required for the Associate/Disassociate operations happening in the background.
- Allow Add New: this configures whether you want to include an "add new" button inside the dropdown tied to the lookup.
- Label (optional): This allows you to insert your own label value. Subgrid labels are slightly larger and semibold, so to keep a uniform look and feel, this option allows you to insert a more "field-like" label.
- Label Location (optional): This allows you to designate the custom label (if provided in the param above) as appearing above the field or beside it.
- Label width (optional): This parameter allows you to set a width to the label space. Useful for custom labels in the "beside" position. Dataverse MDAs use different label widths for different column sizing and aspect ratios. Configure this in "px" or "%" to fit your form layout.
- Custom Filter: This parameter allows injection of a custom odata filter expression to narrow the possible options (e.g. "?statecode eq 0").
- Custom Order: This parameter allows injection of a custom odata order expression to sort the possible options (e.g. "name asc").
- search Mode: This parameter allows you to configure the control for different search 'modes'. It allows for 'Simple' and 'Advanced' search modes. Simple search mode pulls the top 5k rows from the target table during init, then does filtering in-memory. This is a fast, simple UX for small-table searching. Advanced search mode enables the use of Dataverse Search to scan the table with relevance stacking and other configurable options.
- Match Terms: (only applicable if Advanced Search is enabled) This parameter allows you to designate whether multiple search terms in a string should be combined as "All Words" or "Any Word" in the search string.
- Enable Best Effort Searching: (only applicable if Advanced Search is enabled) This parameter allows you to designate whether the control should attempt to search for a term even if it is not found in the search index. This can be useful for partial matches or misspellings.

 ## Limitations
This control is still under active code development. The latest release has added more robust search capabilities, but the code is still an active development project. Known issues and limitations will be listed here as they are identified.

- In Advanced Search mode, when focus leaves the control and then returns, the search results from the last search are still displayed, even though the search string in the box has been cleared. This does not interfere with use of the control, but is a visual bug.

 ## Contributing
 Contributions are welcome! If you have code to suggest, submit a PR: if you have ideas to offer, submit an Issue!