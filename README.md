# utl_excel_experimenting-with-the-new-ods-excel-destination
Experimenting with the new ODS EXCEL Destination    Good Programmers write there own code, great ones steal.   This is stolen stuff from all over the internet, thanks all of you.   Special thanks to Chris Hemedinger and Eric Gebhart    1. Has examples of style elements, inline tags, textdecoration,      tagattr, rotated headings, data superscripts, data subscripts,      autofilters, froxen headers, ods text, titles, footnotes,      heading superscripts and heading subscripts.   2. Creates two workbooks with one sheet in each workbook.   3. Shows two methods to add 'DATA' sheets to either workbook.   4. Has a macro to copy sheets from one workbook to another.      The sheets can have graphs, data, formulas.. or combinations      and will be added(cloned) to the second workbook.

    ```  Experimenting with the new ODS EXCEL Destination                                                                                                             ```
    ```                                                                                                                                                               ```
    ```    Good Programmers write there own code, great ones steal.                                                                                                   ```
    ```    This is stolen stuff from all over the internet, thanks all of you.                                                                                        ```
    ```    Special thanks to Chris Hemedinger and Eric Gebhart                                                                                                        ```
    ```                                                                                                                                                               ```
    ```    1. Has examples of style elements, inline tags, textdecoration,                                                                                            ```
    ```       tagattr, rotated headings, data superscripts, data subscripts,                                                                                          ```
    ```       autofilters, froxen headers, ods text, titles, footnotes,                                                                                               ```
    ```       heading superscripts and heading subscripts.                                                                                                            ```
    ```    2. Creates two workbooks with one sheet in each workbook.                                                                                                  ```
    ```    3. Shows two methods to add 'DATA' sheets to either workbook.                                                                                              ```
    ```    4. Has a macro to copy sheets from one workbook to another.                                                                                                ```
    ```       The sheets can have graphs, data, formulas.. or combinations                                                                                            ```
    ```       and will be added(cloned) to the second workbook.                                                                                                       ```
    ```                                                                                                                                                               ```
    ```  %let fyl=c:\top\xls\&pgm._100rpt.xlsx;                                                                                                                       ```
    ```  %utlfkil(&fyl); * delete if exists;                                                                                                                          ```
    ```                                                                                                                                                               ```
    ```  title;                                                                                                                                                       ```
    ```  footnote;                                                                                                                                                    ```
    ```                                                                                                                                                               ```
    ```  options orientation=landscape nocenter;                                                                                                                      ```
    ```                                                                                                                                                               ```
    ```  *   *   ***   *****  *****   ***                                                                                                                             ```
    ```  **  *  *   *    *    *      *   *                                                                                                                            ```
    ```  * * *  *   *    *    *       *                                                                                                                               ```
    ```  *  **  *   *    *    ****     *                                                                                                                              ```
    ```  *   *  *   *    *    *         *                                                                                                                             ```
    ```  *   *  *   *    *    *      *   *                                                                                                                            ```
    ```  *   *   ***     *    *****   ***;                                                                                                                            ```
    ```                                                                                                                                                               ```
    ```  Notes: (may not be totally correct - my observations)                                                                                                        ```
    ```                                                                                                                                                               ```
    ```    1. Option "Start_at" only uses last instance. It messes up other settings(autofilter?)?                                                                    ```
    ```       Do not use?                                                                                                                                             ```
    ```    2. Results are sensitive to location of 'ods text' and 'titles' and 'footnotes'. Put titles and footnotes at top?                                          ```
    ```    3. Cell with in percentages sort of works but no way to set overall width so percetages are useless?                                                       ```
    ```    4. Absolute_column_width is not needed. Better to use cell_width?                                                                                          ```
    ```    5. I think SAS is creating a zipped IML file ala newer xlsx files?                                                                                         ```
    ```    6. tagattr not needed as often because sas formats, inline tags and style statements are honored?                                                          ```
    ```    7. Need to nest {{{}}} to get all inline formatting to be honored.                                                                                         ```
    ```    8. SAS only supports adding data sheets to an existing workbook?                                                                                           ```
    ```    9. Macro included that uses powershell to copy a sheet from one workbook                                                                                   ```
    ```       to another workbook. Really wanted to use R but could not figure out how.                                                                               ```
    ```       Python can do it cross platform. Unfortunately I fell back on windows only powershell,                                                                  ```
    ```       Afte rall Gates only have about 100 billion dollars.                                                                                                    ```
    ```   10. Because ods excel seems to honor SAS formats, I suspect tagattr may not be as useful.                                                                   ```
    ```       Howver it is useful for totating headers. style(header)={tagattr='rotate:45'} - see below                                                               ```
    ```   11. Also if proc report is your preferred reporting tool then user templates may not be needed.                                                             ```
    ```                                                                                                                                                               ```
    ```  ****    ***    ***   *   *          ***   *   *  *****                                                                                                       ```
    ```   *  *  *   *  *   *  *  *          *   *  **  *  *                                                                                                           ```
    ```   *  *  *   *  *   *  * *           *   *  * * *  *                                                                                                           ```
    ```   ***   *   *  *   *  **            *   *  *  **  ****                                                                                                        ```
    ```   *  *  *   *  *   *  * *           *   *  *   *  *                                                                                                           ```
    ```   *  *  *   *  *   *  *  *          *   *  *   *  *                                                                                                           ```
    ```  ****    ***    ***   *   *          ***   *   *  *****;                                                                                                      ```
    ```                                                                                                                                                               ```
    ```  ods excel file="&fyl" style=pearl                                                                                                                            ```
    ```     options                                                                                                                                                   ```
    ```         (                                                                                                                                                     ```
    ```       /*  start_at                   = "D3"    messes up autofilter? and other stuff */                                                                       ```
    ```           tab_color                  = "red"                                                                                                                  ```
    ```           autofilter                 = 'yes'                                                                                                                  ```
    ```           orientation                = 'landscape'                                                                                                            ```
    ```           zoom                       = "80"                                                                                                                   ```
    ```           suppress_bylines           = 'no'                                                                                                                   ```
    ```           embedded_titles            = 'yes'                                                                                                                  ```
    ```           embedded_footnotes         = 'yes'                                                                                                                  ```
    ```           embed_titles_once          = 'yes'                                                                                                                  ```
    ```           gridlines                  = 'yes'                                                                                                                  ```
    ```           frozen_headers             = 'Yes'                                                                                                                  ```
    ```      /*   absolute_column_width      =  "30pct,22pct,22pct,23pct" not needed */                                                                               ```
    ```           frozen_rowheaders          = 'yes'                                                                                                                  ```
    ```          );                                                                                                                                                   ```
    ```  ;run;quit;                                                                                                                                                   ```
    ```                                                                                                                                                               ```
    ```  ods excel options(sheet_name="utl_100rpt" sheet_interval="none");                                                                                            ```
    ```  ods escapechar='^';                                                                                                                                          ```
    ```                                                                                                                                                               ```
    ```  /*                                                                                                                                                           ```
    ```  I do not see any need for turnining on protect special characters                                                                                            ```
    ```    since we have escape characters for the five below                                                                                                         ```
    ```                                                                                                                                                               ```
    ```  Only 5 chars that need escaping for ods excel?                                                                                                               ```
    ```                                                                                                                                                               ```
    ```  "   &quot;                                                                                                                                                   ```
    ```  '   &apos;                                                                                                                                                   ```
    ```  <   &lt;                                                                                                                                                     ```
    ```  >   &gt;                                                                                                                                                     ```
    ```  &   &amp;                                                                                                                                                    ```
    ```                                                                                                                                                               ```
    ```                                                                                                                                                               ```
    ```  These seem to work                                                                                                                                           ```
    ```                                                                                                                                                               ```
    ```  ~{dagger}                                                                                                                                                    ```
    ```  ~{sigma}                                                                                                                                                     ```
    ```  ~{super text}                                                                                                                                                ```
    ```  ~{sub text}                                                                                                                                                  ```
    ```  ~{raw type text}                                                                                                                                             ```
    ```  ~{style <style><[attributes]>}                                                                                                                               ```
    ```  ~{nbspace count}                                                                                                                                             ```
    ```  ~{newline count}                                                                                                                                             ```
    ```  ~{unicode <Hex|name>}                                                                                                                                        ```
    ```                                                                                                                                                               ```
    ```  style={just=right tagattr='format:##0.###0'};                                                                                                                ```
    ```  style={just=right tagattr='format:$##0.###0'};                                                                                                               ```
    ```                                                                                                                                                               ```
    ```  title3 'Example of ^{nbspace 3} Non-Breaking Spaces Function';                                                                                               ```
    ```  title4 'Example of ^{newline 2} Newline Function';                                                                                                           ```
    ```  title5 'Example of ^{raw \cf12 RAW} RAW function';                                                                                                           ```
    ```  title6 'Example of ^{unicode 03B1} UNICODE function';                                                                                                        ```
    ```                                                                                                                                                               ```
    ```  title "Examples of Functions";                                                                                                                               ```
    ```  title2 'This is ^{style [color=red] Red}';                                                                                                                   ```
    ```  */                                                                                                                                                           ```
    ```                                                                                                                                                               ```
    ```                                                                                                                                                               ```
    ```  * stack region and division in one cell;                                                                                                                     ```
    ```  data UsaPre;                                                                                                                                                 ```
    ```    length stackregdiv $44 statename $64;                                                                                                                      ```
    ```    set sashelp.us_data;                                                                                                                                       ```
    ```    stackregdiv=catx('^{newline}',region,division);                                                                                                            ```
    ```    if statename='Illinois' then statename=cats(statename,'^{super 7}');                                                                                       ```
    ```    if statename='Ohio'     then statename=cats('^S={textdecoration=Underline}',statename);                                                                    ```
    ```    if statename='Indiana'  then statename="^S={foreground=red}^"!!statename;                                                                                  ```
    ```    /*                                                                                                                                                         ```
    ```    does not work                                                                                                                                              ```
    ```    if statename='Ohio'      then   statename=cats('^{font_weight=bold font_size=13pt font_style=italic}^',statename);                                         ```
    ```    */                                                                                                                                                         ```
    ```  run;quit;                                                                                                                                                    ```
    ```                                                                                                                                                               ```
    ```                                                                                                                                                               ```
    ```  proc sort data=UsaPre out=UsaSrt noequals;                                                                                                                   ```
    ```  by REGION                                                                                                                                                    ```
    ```     DIVISION                                                                                                                                                  ```
    ```     STATENAME                                                                                                                                                 ```
    ```  ;                                                                                                                                                            ```
    ```  run;quit;                                                                                                                                                    ```
    ```                                                                                                                                                               ```
    ```  * highlight certain sells;                                                                                                                                   ```
    ```  proc format;                                                                                                                                                 ```
    ```     value cback low-25   ='light red'                                                                                                                         ```
    ```                  26-32   = 'yellow'                                                                                                                           ```
    ```                  33-high = 'white';                                                                                                                           ```
    ```  run;                                                                                                                                                         ```
    ```                                                                                                                                                               ```
    ```  title;footnote;                                                                                                                                              ```
    ```                                                                                                                                                               ```
    ```  title1 justify=center h=15pt "Experimenting with the ODS Excel destination with undocumented options?";                                                      ```
    ```  title2 justify=center bspace=0cm h=15pt "&sysdate Roger DeAngelis";                                                                                          ```
    ```  title3 h=15pt "Example of ^{super ^{style [foreground=red] red ^{style [foreground=green]green } and ^{style [foreground=blue] blue}}} formatting";          ```
    ```  title4 h=15pt "Example of ^{dagger ^{style [foreground=red] red ^{style [foreground=green]green } and ^{style [foreground=blue] blue}}}                      ```
    ```     ^{sigma ^{style [foreground=red] red ^{style [foreground=green]green } and ^{style [foreground=blue] blue}}} formatting";                                 ```
    ```  footnote1 justify=left bspace=0cm h=15pt "SAS Footnote1";                                                                                                    ```
    ```  footnote2 justify=left bspace=0cm h=15pt "SAS Footnote2";                                                                                                    ```
    ```  run;quit;                                                                                                                                                    ```
    ```                                                                                                                                                               ```
    ```  ods text="^S={font_size=14pt just=r outputwidth=2400px font_weight=bold color=blue font_face=arial}^United States Census Statistics from 1910 to 2010";      ```
    ```  ods text="^S={font_size=14pt font_face=arial color=brown}^Safety Set";                                                                                       ```
    ```  run;quit;                                                                                                                                                    ```
    ```                                                                                                                                                               ```
    ```  PROC REPORT DATA=UsaSrt SPLIT="#" nowd                                                                                                                       ```
    ```     style(header)={just=center font_weight=bold color=green font_size=13pt}                                                                                   ```
    ```     style(column)={font_size=11pt font_face=arial};                                                                                                           ```
    ```  COLUMN                                                                                                                                                       ```
    ```    ( "Census data on Us Population from 1910 to 2010"                                                                                                         ```
    ```     ( "Geographic Variables"                                                                                                                                  ```
    ```     REGION                                                                                                                                                    ```
    ```     DIVISION                                                                                                                                                  ```
    ```     STACKREGDIV                                                                                                                                               ```
    ```     STATENAME                                                                                                                                                 ```
    ```     STATE                                                                                                                                                     ```
    ```     STATECODE                                                                                                                                                 ```
    ```     )                                                                                                                                                         ```
    ```     ( "Population per Square Mile from Decenial US Census data ad Median Statistics by Reagion and Division"                                                  ```
    ```     DENSITY_1910                                                                                                                                              ```
    ```     DENSITY_1920                                                                                                                                              ```
    ```     DENSITY_1930                                                                                                                                              ```
    ```     DENSITY_1940                                                                                                                                              ```
    ```     DENSITY_1950                                                                                                                                              ```
    ```     DENSITY_1960                                                                                                                                              ```
    ```     DENSITY_1970                                                                                                                                              ```
    ```     DENSITY_1980                                                                                                                                              ```
    ```     DENSITY_1990                                                                                                                                              ```
    ```     DENSITY_2000                                                                                                                                              ```
    ```     DENSITY_2010                                                                                                                                              ```
    ```     )                                                                                                                                                         ```
    ```    );                                                                                                                                                         ```
    ```                                                                                                                                                               ```
    ```  DEFINE  REGION       / order FORMAT= $9.    style(column)={cellwidth=1in}    "US#Regions"  style(header)={tagattr='rotate:45'};                              ```
    ```  DEFINE  DIVISION     / order FORMAT= $18.   style(column)={just=l cellwidth=1.4in}  style(header)={just=left color=red}" US";                                ```
    ```  DEFINE  STACKREGDIV  / order FORMAT= $44.   style(column)={just=l cellwidth=1.4in}  "Stacked#Region#Division" ;                                              ```
    ```  DEFINE  STATENAME    / order FORMAT= $64.   style(column)={just=l cellwidth=1.4in}  "Name of State#or Region";                                               ```
    ```  DEFINE  STATE        / order FORMAT= BEST9. style(column)={just=c cellwidth=.5in}    "State#Fips#Code" ;                                                     ```
    ```  DEFINE  STATECODE    / order FORMAT= $2.    style(column)={just=c cellwidth=.5in}   "State#Code";                                                            ```
    ```  DEFINE  DENSITY_1910 / median FORMAT= COMMA8.1  "1910 People per#Square Mile" style(column)={just=r cellwidth=1in background=cback.};                        ```
    ```  DEFINE  DENSITY_1920 / median FORMAT= COMMA8.1  "1920 People per#Square Mile" style(column)={just=r cellwidth=1in};                                          ```
    ```  DEFINE  DENSITY_1930 / median FORMAT= COMMA8.1  "1930 People per#Square Mile" style(column)={just=r cellwidth=1in};                                          ```
    ```  DEFINE  DENSITY_1940 / median FORMAT= COMMA8.1  "1940 People per#Square Mile" style(column)={just=r cellwidth=1in};                                          ```
    ```  DEFINE  DENSITY_1950 / median FORMAT= COMMA8.1  "1950 People per#Square Mile" style(column)={just=r cellwidth=1in};                                          ```
    ```  DEFINE  DENSITY_1960 / median FORMAT= COMMA8.1  "1960 People per#Square Mile" style(column)={just=r cellwidth=1in};                                          ```
    ```  DEFINE  DENSITY_1970 / median FORMAT= COMMA8.1  "1970 People per#Square Mile" style(column)={just=r cellwidth=1in};                                          ```
    ```  DEFINE  DENSITY_1980 / median FORMAT= COMMA8.1  "1980 People per#Square Mile" style(column)={just=r cellwidth=1in};                                          ```
    ```  DEFINE  DENSITY_1990 / median FORMAT= COMMA8.1  "1990 People per#Square Mile" style(column)={just=r cellwidth=1in};                                          ```
    ```  DEFINE  DENSITY_2000 / median FORMAT= COMMA8.1  "2000 People per#Square Mile" style(column)={just=r cellwidth=1in};                                          ```
    ```  DEFINE  DENSITY_2010 / median FORMAT= COMMA8.1  "2010 People per#Square Mile" style(column)={just=r cellwidth=1in};                                          ```
    ```  break after region/summarize;                                                                                                                                ```
    ```  compute after region;                                                                                                                                        ```
    ```    statename="Median";                                                                                                                                        ```
    ```  endcomp;                                                                                                                                                     ```
    ```  break after division/summarize;                                                                                                                              ```
    ```  compute after/style={just=left font_size=11pt font_weight=bold};                                                                                             ```
    ```    lyn="^{newline}Compute after reort table Experimenting with the ODS Excel destination";                                                                    ```
    ```    line lyn $96.;                                                                                                                                             ```
    ```  endcomp;                                                                                                                                                     ```
    ```  run;quit;                                                                                                                                                    ```
    ```  ods  text='^S={font_size=12pt font_weight=bold color=brown}^Page 1 of 1';                                                                                    ```
    ```  ods  text='^S={font_size=10pt font_weight=bold font_style=italic}^post report text1}';                                                                       ```
    ```  run;quit;                                                                                                                                                    ```
    ```                                                                                                                                                               ```
    ```  ods excel style=minimal;                                                                                                                                     ```
    ```  title;footnote;run;quit;                                                                                                                                     ```
    ```  ods graphics / height=800px width=1100px noborder;                                                                                                           ```
    ```  title1 h=20pt "MSRP for Cars";                                                                                                                               ```
    ```  proc sgplot data=sashelp.cars;                                                                                                                               ```
    ```  histogram msrp;                                                                                                                                              ```
    ```  run;                                                                                                                                                         ```
    ```                                                                                                                                                               ```
    ```  ods excel close;                                                                                                                                             ```
    ```                                                                                                                                                               ```
    ```                                                                                                                                                               ```
    ```  ****    ***    ***   *   *   ***                                                                                                                             ```
    ```   *  *  *   *  *   *  *  *   *   *                                                                                                                            ```
    ```   *  *  *   *  *   *  * *        *                                                                                                                            ```
    ```   ***   *   *  *   *  **       **                                                                                                                             ```
    ```   *  *  *   *  *   *  * *     *                                                                                                                               ```
    ```   *  *  *   *  *   *  *  *   *                                                                                                                                ```
    ```  ****    ***    ***   *   *  *****;                                                                                                                           ```
    ```                                                                                                                                                               ```
    ```                                                                                                                                                               ```
    ```  %let fyl=c:\top\xls\&pgm._200rpt.xlsx;                                                                                                                       ```
    ```                                                                                                                                                               ```
    ```  %utlfkil(&fyl); * delete file;                                                                                                                               ```
    ```                                                                                                                                                               ```
    ```  title;                                                                                                                                                       ```
    ```  footnote;                                                                                                                                                    ```
    ```                                                                                                                                                               ```
    ```  options orientation=landscape nocenter;                                                                                                                      ```
    ```                                                                                                                                                               ```
    ```  filename out "&fyl";                                                                                                                                         ```
    ```                                                                                                                                                               ```
    ```  ods excel file="&fyl" style=pearl                                                                                                                            ```
    ```     options                                                                                                                                                   ```
    ```         (                                                                                                                                                     ```
    ```       /*  start_at                   = "D3"    messes up autofilter? and other stuff */                                                                       ```
    ```           tab_color                  = "red"                                                                                                                  ```
    ```           autofilter                 = 'yes'                                                                                                                  ```
    ```           orientation                = 'landscape'                                                                                                            ```
    ```           zoom                       = "80"                                                                                                                   ```
    ```           suppress_bylines           = 'no'                                                                                                                   ```
    ```           embedded_titles            = 'yes'                                                                                                                  ```
    ```           embedded_footnotes         = 'yes'                                                                                                                  ```
    ```           embed_titles_once          = 'yes'                                                                                                                  ```
    ```           gridlines                  = 'yes'                                                                                                                  ```
    ```           frozen_headers             = 'Yes'                                                                                                                  ```
    ```      /*   absolute_column_width      =  "30pct,22pct,22pct,23pct" not needed */                                                                               ```
    ```           frozen_rowheaders          = 'yes'                                                                                                                  ```
    ```          );                                                                                                                                                   ```
    ```  ;run;quit;                                                                                                                                                   ```
    ```                                                                                                                                                               ```
    ```  ods excel options(sheet_name="utl_200rpt" sheet_interval="none");                                                                                            ```
    ```  ods escapechar='^';                                                                                                                                          ```
    ```                                                                                                                                                               ```
    ```  title1 h=20pt "SAS Cars Dataset";                                                                                                                            ```
    ```  title2 h=15pt "Add Sheet utl_200rpt.xlsx to previous Workbook";                                                                                              ```
    ```                                                                                                                                                               ```
    ```  PROC REPORT DATA=SASHELP.CARS LS=171 PS=65  SPLIT="/" WRAP NOCENTER MISSING ;                                                                                ```
    ```  COLUMN  MAKE MODEL TYPE ORIGIN DRIVETRAIN MSRP INVOICE ENGINESIZE                                                                                            ```
    ```  CYLINDERS HORSEPOWER MPG_CITY MPG_HIGHWAY WEIGHT WHEELBASE LENGTH;                                                                                           ```
    ```                                                                                                                                                               ```
    ```  DEFINE  MAKE / DISPLAY       style={cellwidth=.8in}     LEFT "MAKE" ;                                                                                        ```
    ```  DEFINE  MODEL / DISPLAY      style={cellwidth=1.5in}    LEFT "MODEL" ;                                                                                       ```
    ```  DEFINE  TYPE / DISPLAY       style={cellwidth=.6in}     LEFT "TYPE" ;                                                                                        ```
    ```  DEFINE  ORIGIN / DISPLAY     style={cellwidth=.5in}     LEFT "ORIGIN" ;                                                                                      ```
    ```  DEFINE  DRIVETRAIN / DISPLAY style={cellwidth=.6in}     LEFT "DRIVETRAIN" ;                                                                                  ```
    ```  DEFINE  MSRP / SUM           style={cellwidth=.6in}     RIGHT "MSRP" ;                                                                                       ```
    ```  DEFINE  INVOICE / SUM        style={cellwidth=.6in}     RIGHT "INVOICE" ;                                                                                    ```
    ```  DEFINE  ENGINESIZE / SUM     style={cellwidth=.6in}     RIGHT "Engine Size (L)" ;                                                                            ```
    ```  DEFINE  CYLINDERS / SUM      style={cellwidth=.6in}     RIGHT "CYLINDERS" ;                                                                                  ```
    ```  DEFINE  HORSEPOWER / SUM     style={cellwidth=.6in}     RIGHT "HORSEPOWER" ;                                                                                 ```
    ```  DEFINE  MPG_CITY / SUM       style={cellwidth=.6in}     RIGHT "MPG (City)" ;                                                                                 ```
    ```  DEFINE  MPG_HIGHWAY / SUM    style={cellwidth=.6in}     RIGHT "MPG (Highway)" ;                                                                              ```
    ```  DEFINE  WEIGHT / SUM         style={cellwidth=.6in}     RIGHT "Weight (LBS)" ;                                                                               ```
    ```  DEFINE  WHEELBASE / SUM      style={cellwidth=.6in}     RIGHT "Wheelbase (IN)" ;                                                                             ```
    ```  DEFINE  LENGTH / SUM         style={cellwidth=.6in}     RIGHT "Length (IN)" ;                                                                                ```
    ```  run;quit;                                                                                                                                                    ```
    ```  ods excel close;                                                                                                                                             ```
    ```                                                                                                                                                               ```
    ```    *    ****   ****           ***   *   *  *****  *****  *****   ***                                                                                          ```
    ```   * *    *  *   *  *         *   *  *   *  *      *        *    *   *                                                                                         ```
    ```  *   *   *  *   *  *          *     *   *  *      *        *     *                                                                                            ```
    ```  *****   *  *   *  *           *    *****  ****   ****     *      *                                                                                           ```
    ```  *   *   *  *   *  *            *   *   *  *      *        *       *                                                                                          ```
    ```  *   *   *  *   *  *         *   *  *   *  *      *        *    *   *                                                                                         ```
    ```  *   *  ****   ****           ***   *   *  *****  *****    *     ***;                                                                                         ```
    ```                                                                                                                                                               ```
    ```  * ADD DATA ONLY USING EXPORT;                                                                                                                                ```
    ```  %let fyl=c:\top\xls\&pgm._200rpt.xlsx;                                                                                                                       ```
    ```  filename out "&fyl";                                                                                                                                         ```
    ```  proc export data=sashelp.class                                                                                                                               ```
    ```    dbms=xlsx                                                                                                                                                  ```
    ```    outfile=out replace;                                                                                                                                       ```
    ```    sheet="Friend Details";                                                                                                                                    ```
    ```  run;                                                                                                                                                         ```
    ```  filename out clear;                                                                                                                                          ```
    ```  run;quit;                                                                                                                                                    ```
    ```                                                                                                                                                               ```
    ```  * ADD DATA ONLY USING LIBNAME;                                                                                                                               ```
    ```  libname xls "&fyl";                                                                                                                                          ```
    ```  data xls.addthis;                                                                                                                                            ```
    ```    set sashelp.class;                                                                                                                                         ```
    ```  run;quit;                                                                                                                                                    ```
    ```  libname xls clear;                                                                                                                                           ```
    ```                                                                                                                                                               ```
    ```   ***    ***   ****   *   *          ***   *   *  *****  *****  *****   ***                                                                                   ```
    ```  *   *  *   *  *   *  *   *         *   *  *   *  *      *        *    *   *                                                                                  ```
    ```  *      *   *  *   *   * *           *     *   *  *      *        *     *                                                                                     ```
    ```  *      *   *  ****     *             *    *****  ****   ****     *      *                                                                                    ```
    ```  *      *   *  *        *              *   *   *  *      *        *       *                                                                                   ```
    ```  *   *  *   *  *        *           *   *  *   *  *      *        *    *   *                                                                                  ```
    ```   ***    ***   *        *            ***   *   *  *****  *****    *     ***;                                                                                  ```
    ```                                                                                                                                                               ```
    ```  * COPY A SHEET FROM WORKBOOK1 TO WORKBOOK2;                                                                                                                  ```
    ```                                                                                                                                                               ```
    ```  %utl_copysheet(                                                                                                                                              ```
    ```     frombook=c:\top\xls\&pgm._100rpt.xlsx                                                                                                                     ```
    ```    ,fromsheet=utl_200rpt                                                                                                                                      ```
    ```                                                                                                                                                               ```
    ```    ,tobook=c:\top\xls\&pgm._200rpt.xlsx                                                                                                                       ```
    ```    ,tosheet=utl_100rpt                                                                                                                                        ```
    ```    );                                                                                                                                                         ```
    ```                                                                                                                                                               ```
    ```  *   *    *     ***   ****    ***    ***                                                                                                                      ```
    ```  ** **   * *   *   *  *   *  *   *  *   *                                                                                                                     ```
    ```  * * *  *   *  *      *   *  *   *   *                                                                                                                        ```
    ```  *   *  *****  *      ****   *   *    *                                                                                                                       ```
    ```  *   *  *   *  *      * *    *   *     *                                                                                                                      ```
    ```  *   *  *   *  *   *  *  *   *   *  *   *                                                                                                                     ```
    ```  *   *  *   *   ***   *   *   ***    ***;                                                                                                                     ```
    ```                                                                                                                                                               ```
    ```  options lrecl=32756; * fixes trucation with cards in 9.3;                                                                                                    ```
    ```  * mac var _o is autocall folder;                                                                                                                             ```
    ```  data _null_;file "c:\oto\utl_copysheet.sas";input;put _infile_;putlog _infile_;                                                                              ```
    ```  cards4;                                                                                                                                                      ```
    ```  %macro utl_copysheet(                                                                                                                                        ```
    ```          frombook=c:\top\xls\&pgm._100rpt.xlsx                                                                                                                ```
    ```         ,tobook=c:\top\xls\&pgm._200rpt.xlsx                                                                                                                  ```
    ```         ,fromsheet=utl_200rpt                                                                                                                                 ```
    ```         ,tosheet=utl_100rpt                                                                                                                                   ```
    ```         )/ des="Copy a sheet from one workbook to the another workbook";                                                                                      ```
    ```                                                                                                                                                               ```
    ```      %local __cmd;                                                                                                                                            ```
    ```                                                                                                                                                               ```
    ```      /*                                                                                                                                                       ```
    ```        For testing without macro call                                                                                                                         ```
    ```        %let frombook=c:\top\xls\&pgm._100rpt.xlsx;                                                                                                            ```
    ```        %let tobook=c:\top\xls\&pgm._200rpt.xlsx;                                                                                                              ```
    ```        %let fromsheet=utl_200rpt;                                                                                                                             ```
    ```        %let tosheet=utl_100rpt;                                                                                                                               ```
    ```      */                                                                                                                                                       ```
    ```                                                                                                                                                               ```
    ```      proc sql;                                                                                                                                                ```
    ```       create                                                                                                                                                  ```
    ```         table __utl_copysheet (chr char(80));insert into __utl_copysheet                                                                                      ```
    ```      VALUES("$file1 = '&frombook' # source's fullpath                               ")                                                                        ```
    ```      VALUES("$file2 = '&tobook' # destination's fullpath                            ")                                                                        ```
    ```      VALUES("$xl = new-object -c excel.application                                  ")                                                                        ```
    ```      VALUES("$xl.displayAlerts = $false # don't prompt the user                     ")                                                                        ```
    ```      VALUES("$wb2 = $xl.workbooks.open($file1, $null, $true) # open source, readonly")                                                                        ```
    ```      VALUES("$wb1 = $xl.workbooks.open($file2) # open target                        ")                                                                        ```
    ```      VALUES("$sh1_wb1 = $wb1.sheets.item('&fromsheet') # 2nd sheet in destination    ")                                                                       ```
    ```      VALUES("$sheetToCopy = $wb2.sheets.item('&tosheet') # source sheet to copy   ")                                                                          ```
    ```      VALUES("$sheetToCopy.copy($sh1_wb1) # copy source sheet to destination workbook")                                                                        ```
    ```      VALUES("$wb2.close($false) # close source workbook w/o saving                  ")                                                                        ```
    ```      VALUES("$wb1.close($true) # close and save destination workbook                ")                                                                        ```
    ```      VALUES("$xl.quit()                                                             ")                                                                        ```
    ```      VALUES("spps -n excel                                                          ")                                                                        ```
    ```      ;quit;                                                                                                                                                   ```
    ```                                                                                                                                                               ```
    ```      %utlfkil(%sysfunc(pathname(work))\ps1.ps1);                                                                                                              ```
    ```                                                                                                                                                               ```
    ```      filename _ps1 "%sysfunc(pathname(work))\ps1.ps1";                                                                                                        ```
    ```      data _null_;                                                                                                                                             ```
    ```        file _ps1;                                                                                                                                             ```
    ```        set __utl_copysheet;                                                                                                                                   ```
    ```        put chr;                                                                                                                                               ```
    ```        putlog chr;                                                                                                                                            ```
    ```        if _n_=1 then do;                                                                                                                                      ```
    ```          cmd=catx(' ',"'powershell -Command",cats('"',"%sysfunc(pathname(work))\ps1.ps1",cats('"',"'")));                                                     ```
    ```          putlog cmd=;                                                                                                                                         ```
    ```          call symputx('__cmd',cmd);                                                                                                                           ```
    ```        end;                                                                                                                                                   ```
    ```      run;quit;                                                                                                                                                ```
    ```      ;;;;                                                                                                                                                     ```
    ```      quit;                                                                                                                                                    ```
    ```      filename _ps1 clear;                                                                                                                                     ```
    ```      run;quit;                                                                                                                                                ```
    ```                                                                                                                                                               ```
    ```      options xwait xsync;run;quit;                                                                                                                            ```
    ```      * you can paste this into a dos window for testing;                                                                                                      ```
    ```      systask kill _ps1;systask command &__cmd taskname=_ps1;                                                                                                  ```
    ```      waitfor _ps1;                                                                                                                                            ```
    ```                                                                                                                                                               ```
    ```  %mend utl_copysheet;                                                                                                                                         ```
    ```  ;;;;                                                                                                                                                         ```
    ```  run;quit;                                                                                                                                                    ```
    ```                                                                                                                                                               ```
    ```  %inc "c:\oto\utl_copysheet.sas";                                                                                                                             ```
    ```                                                                                                                                                               ```
    ```  * in case you have been testing interactively;                                                                                                               ```
    ```  %symdel fromsheet;                                                                                                                                           ```
    ```  %symdel tosheet;                                                                                                                                             ```
    ```  %symdel frombook;                                                                                                                                            ```
    ```  %symdel tobook;                                                                                                                                              ```
    ```  %utl_copysheet;                                                                                                                                              ```
    ```                                                                                                                                                               ```
    ```  *   *  *****  *      *****  *   *  *****  *                                                                                                                  ```
    ```  *   *    *    *      *      *  *     *    *                                                                                                                  ```
    ```  *   *    *    *      *      * *      *    *                                                                                                                  ```
    ```  *   *    *    *      ****   **       *    *                                                                                                                  ```
    ```  *   *    *    *      *      * *      *    *                                                                                                                  ```
    ```  *   *    *    *      *      *  *     *    *                                                                                                                  ```
    ```   ***     *    *****  *      *   *  *****  *****;                                                                                                             ```
    ```                                                                                                                                                               ```
    ```  %macro utlfkil                                                                                                                                               ```
    ```      (                                                                                                                                                        ```
    ```      utlfkil                                                                                                                                                  ```
    ```      ) / des="delete an external file";                                                                                                                       ```
    ```                                                                                                                                                               ```
    ```                                                                                                                                                               ```
    ```      /*-------------------------------------------------*\                                                                                                    ```
    ```      |                                                   |                                                                                                    ```
    ```      |  Delete an external file                          |                                                                                                    ```
    ```      |   From SAS macro guide                                                |                                                                                ```
    ```      |  Sample invocations                               |                                                                                                    ```
    ```      |                                                   |                                                                                                    ```
    ```      |  WIN95                                            |                                                                                                    ```
    ```      |  %utlfkil(c:\dat\utlfkil.sas);                    |                                                                                                    ```
    ```      |                                                   |                                                                                                    ```
    ```      |                                                   |                                                                                                    ```
    ```      |  Solaris 2.5                                      |                                                                                                    ```
    ```      |  %utlfkil(/home/deangel/delete.dat);              |                                                                                                    ```
    ```      |                                                   |                                                                                                    ```
    ```      |                                                   |                                                                                                    ```
    ```      |  Roger DeAngelis                                  |                                                                                                    ```
    ```      |                                                   |                                                                                                    ```
    ```      \*-------------------------------------------------*/                                                                                                    ```
    ```                                                                                                                                                               ```
    ```      %local urc;                                                                                                                                              ```
    ```                                                                                                                                                               ```
    ```      /*-------------------------------------------------*\                                                                                                    ```
    ```      | Open file   -- assign file reference              |                                                                                                    ```
    ```      \*-------------------------------------------------*/                                                                                                    ```
    ```                                                                                                                                                               ```
    ```      %let urc = %sysfunc(filename(fname,%quote(&utlfkil)));                                                                                                   ```
    ```                                                                                                                                                               ```
    ```      /*-------------------------------------------------*\                                                                                                    ```
    ```      | Delete file if it exits                           |                                                                                                    ```
    ```      \*-------------------------------------------------*/                                                                                                    ```
    ```                                                                                                                                                               ```
    ```      %if &urc = 0 and %sysfunc(fexist(&fname)) %then                                                                                                          ```
    ```          %let urc = %sysfunc(fdelete(&fname));                                                                                                                ```
    ```                                                                                                                                                               ```
    ```      /*-------------------------------------------------*\                                                                                                    ```
    ```      | Close file  -- deassign file reference            |                                                                                                    ```
    ```      \*-------------------------------------------------*/                                                                                                    ```
    ```                                                                                                                                                               ```
    ```      %let urc = %sysfunc(filename(fname,''));                                                                                                                 ```
    ```                                                                                                                                                               ```
    ```    run;                                                                                                                                                       ```
    ```                                                                                                                                                               ```
    ```  %mend utlfkil;                                                                                                                                               ```

