# Overview

This was a fun demo in the usage of a UserForm along with VBA to produce an dynamic event based navigation menu. 


## Operation
- Right now, as this is a demo, the values that populate the navigation menu come from `Sheet2` in the Workbook. 
- The cells that the values for the navigation menu are `Range("A1:B15")`
- In `clsSection_NavigationContainer.cls`, you can change the data origination location.
    - How it currently looks `permissionsArray = Worksheets("Sheet2").Range("A1:B15").value`
- When clicking on the top level, only the initial click is recorded. Any successive clicks to top level navigation items won't trigger events. 
    - This is to allow users to move around without navigating.
- Any time that the user clicks a sub-nav item, an event till fire. 
- To see the debug usage of this, add the **immediate window**.

## Thoughts
- As this is just a demo, I have not included code to make it fully dynamic. If I were to do that, I'be be using a database to pull in the values -- I'll demo this later on.
- In the future:
    - I'll add a scrollbar to handle any overflows. 
    - I'll add how to make the body of the UserForm `reactive` and dynamically populate **sections** / **pages** based on the item that was clicked in the navigation


## Visual Demo
![](https://github.com/Brostoffed/excel-vba-navigation-menu/blob/main/assets/Excel-VBA-Navigation-Dropdown-Demo-480p.gif)