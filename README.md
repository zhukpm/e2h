# Excel2Html converter (2017) #

A small, simple and configurable library for converting Excel-sheets (MS Excel) to html-tables using Apache POI project. Primarily
developed to publish simple (and complex) Excel-reports to corporate portals.

This converter converts a cell range on a sheet to standalone html `<table> ... </table>` inserting `style` attributes.

## Installation ##

1. Clone this repository using `git`
2. Install via the command line

      `mvn clean install`

3. Add as a dependency to your project in `pom.xml`:

```xml
    <dependency>
        <groupId>ru.perrymason</groupId>
        <artifactId>e2h</artifactId>
        <version>0.1.0</version>
    </dependency>
```

## Use cases ##

### Converting Excel file ###

There is an `E2HFileFacade` class created for file-to-file converting. It may be used for multiple sheets from the same
Excel file.

```java
    File excelFile = ...
    File report1 = ...
    File report2 = ...

    E2HFileFacade facade = new E2HFileFacade(excelFile);
    facade.addOption(E2HOption.STANDARD_OPTIONS);

    facade.selectSheet("Report 1");
    facade.selectCellRange("C3:K19");
    facade.writeHtml(report1);

    facade.selectSheet("Report 2");
    facade.selectCellRange("B2:G36");
    facade.writeHtml(report2);

    facade.close();
```

### Converting a POI workbook ###

You can use an `Excel2Html` class directly if you want to specify  different options, data formatters and output streams.

```java
    Sheet sheet = ...
    CellRangeAddress range = ...
    OutputStream stream = ...

    Excel2Html excel2Html = Excel2Html.getConverter(sheet, range);
    excel2Html.addOption(E2HOption.FONT);
    excel2Html.addOption(E2HOption.BORDERS);
    excel2Html.writeHtml(stream);

    stream.flush();
    stream.close();
```

## Limitations ##

There are some limitations, including
* complex cell styles;
* data formatting;
* colors;
* etc.

But it kinda works with standard simple Excel reports.


## Issues ##

Feel free to report any issues, proposals, etc. They may be even fixed and implemented later.