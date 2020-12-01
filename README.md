# A Power Apps Component Framework FluentUI Grid

An example PCF component that can be added to a Contact form and then used to display and perform basic filtering on a data set of appointment slots, and then create a booking record against one of the slots.

Written as an exercise to learn the Power Apps Component Framework and FluentUI. A lot of inspiration was taken from [Michal Turzynski's Fluent UI DataList project](https://github.com/michal-turzynski/pcf-fluentui-grid)

Use this command to build the component: `msbuild /p:pcfbuildmode=production`. The pcfbuildmode switch is required as otherwise the WebResource files that are generated may be too large to be imported to a CDS environment.

![Screen Capture](/img/picker.gif)
