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
- Label max width (optional): This parameter allows you to set a max width to the label space. Useful for custom labels, in the "beside" position. Dataverse MDAs use different label widths for different column sizing and aspect ratios. Configure this in "px" or "%" to fit your form layout.
- Custom OData Filter: This parameter allows injection of a custom odata filter expression to narrow the possible options (e.g. "?filter=(statecode eq 0)").

 ## Limitations
 As currently released, this control takes a very simplistic approach to search, retrieving a 'Top 5000' and performing filtering in memory on the client machine. This initial release is about getting the user experience right, and future releases will provide more robust, large-table search versatility. In the meantime, use of this control should be limited to tables with fewer than 5k rows.

 ## Contributing
 Contributions are welcome! Work is underway to support Dataverse Search as a more robust search engine, and I am interested in potentially supporting a Polymorphic experience, but haven't settled on a design for that yet. If you have code to suggest, submit a PR: if you have ideas to offer, submit an Issue!
