%let pgm=utl-linked-trimmed-table-of-contents-in-rtf-with-proc-report-and-msword-cntl-a-f9-and-cntl-click;

%stop_submission; /* for development */

Creating a linked trimmed table of contents in rtf and msword using cntl a f9 and cntl click;

Don't forget to highlight the TOC and hit cntl-a anf f9 in word to fill in the table of contents.
To execute the link put the cursor on any toc entry in the YOC and hit cnth-left-click (cntl left-mouse-button)
To return to the TOC from any page put the cursor on the blue  'RETURN' text and type contrl click.

Only works with proc report?
Only report can remove the excess levels in the table of contents?
Best with classic 1980s sas DMS editor, should work with other sas editors.

Note: The table of contents is compliled using proclabels and proc report contents=" " options,
Cntl-A and f9 executes and embedded  rtf script to populate the TOC.
The tables on contents is trimmed by using report options contents=" "
and break before header / contents=" " page;

RTF
https://tinyurl.com/5h8kn7tj
https://github.com/rogerjdeangelis/utl-linked-trimmed-table-of-contents-in-rtf-with-proc-report-and-msword-cntl-a-f9-and-cntl-click/blob/main/utl_toc.rtf


DOC  (open the rtf file with msword and save as docx(
https://tinyurl.com/5ebcc5dm
https://github.com/rogerjdeangelis/utl-linked-trimmed-table-of-contents-in-rtf-with-proc-report-and-msword-cntl-a-f9-and-cntl-click/blob/main/utl_toc.docx



RELATED RTF REPOS
=================

https://github.com/rogerjdeangelis/utl-removing-unwanted-bookmarks-in-pdf-table-of-contents-toc
https://github.com/rogerjdeangelis/ods_rtf_mutiple_justifications_within_one_compute_block
https://github.com/rogerjdeangelis/utl-adding-images-to-rtf-and-word-docs
https://github.com/rogerjdeangelis/utl-create-a-simple-n-percent-clinical-table-in-r-sas-wps-python-output-pdf-rtf-xlsx-html-list
https://github.com/rogerjdeangelis/utl-example-rtf-excel-and-pdf-reports-using-all-sas-provided-style-templates
https://github.com/rogerjdeangelis/utl-putting-a-frame-around-text-in-doc-rtf-and-pdf-ods-destinations-with-and-without-layout
https://github.com/rogerjdeangelis/utl-retaining-header-row-across-pages-on-ods-rtf-proc-report
https://github.com/rogerjdeangelis/utl-sas-macro-to-combine-rtf-files-into-one-single-file
https://github.com/rogerjdeangelis/utl-sas-ods-underlining-text-in-html-pdf-and-rtf
https://github.com/rogerjdeangelis/utl_different_header_and_footer_for_rtf_document_and_for_table
https://github.com/rogerjdeangelis/utl_dropping-down-to-powershell-and-converting-doc-and-rtf-files-to-pdfs
https://github.com/rogerjdeangelis/utl_ods_pdf_and_rtf_two_different_page_titles_on_the_same_page
https://github.com/rogerjdeangelis/utl_report_does_not_show_group_variable_across_new_pages_in_rtf_and_pdf

/*                   _
(_)_ __  _ __  _   _| |_
| | `_ \| `_ \| | | | __|
| | | | | |_) | |_| | |_
|_|_| |_| .__/ \__,_|\__|
        |_|
*/

You also neeed the folder d:\rtf;

sashelp.class(where=(sex='F'))
sashelp.class(where=(sex='M'))

/*

 _ __  _ __ ___   ___ ___  ___ ___
| `_ \| `__/ _ \ / __/ _ \/ __/ __|
| |_) | | | (_) | (_|  __/\__ \__ \
| .__/|_|  \___/ \___\___||___/___/
|_|
*/


%utlfkil(d:/rtf/utl_toc.rtf);

title;
footnote;

ods listing close;

ods rtf path="d:\rtf" file="utl_toc.rtf" contents toc_data bodytitle style=statistical;

ods escapechar="~";

ods rtf startpage=now; /* for some reason this eliminates blank pages            */
                       /* especially if you have data steps or processes between */
ods proclabel="Females";
proc report data= sashelp.class(where=(sex='F')) contents=" ";
title "~S={just=left font_weight=bold fontsize=16pt} Females";
  col sex name age height weight;
  define sex /group noprint;
  break before sex / contents=" " page;
  compute after / style={font_weight=bold};
   lyn = "~S={URL='utl_toc.rtf#PAGE=1' just=left color=blue } --RETURN--" ;
   line lyn $64.;
  endcomp;
run;quit;
ods rtf startpage=no;  /* for some reason this eliminates blank pages            */

ods rtf startpage=now; /* for some reason this eliminates blank pages            */
ods proclabel="Males";
proc report data= sashelp.class(where=(sex='M')) contents=" ";
title "~S={just=left font_weight=bold font_size=12pt fontsize=16pt} Males";
  col sex name age height weight;
  define sex /group noprint;
  break before sex / contents=" " page;
  compute after / style={font_weight=bold};
   lyn = "~S={URL='utl_toc.rtf#PAGE=1' just=left color=blue } --RETURN--" ;
   line lyn $64.;
  endcomp;
run;quit;
ods rtf startpage=no; /* for some reason this eliminates blank pages */

ods rtf close;
ods listing;

/*           _               _
  ___  _   _| |_ _ __  _   _| |_
 / _ \| | | | __| `_ \| | | | __|
| (_) | |_| | |_| |_) | |_| | |_
 \___/ \__,_|\__| .__/ \__,_|\__|
                |_|
*/
  footnote "~S={URL='utl_toc.rtf#PAGE=1' just=left color=blue } --RETURN--" ;

/**************************************************************************************************************************/
/*                                                                                                                        */
/* NOTES:                                                                                                                 */
/*                                                                                                                        */
/*  The table of contents will not be populated unless you issue the                                                      */
/*  keystrokes cntl-a then F9                                                                                             */
/*                                                                                                                        */
/*  To use the link, place the cursor on Females and issue keystroke cntl-click                                           */
/*  Females  ,............................    1                                                                           */
/*                                                                                                                        */
/*                                                                                                                        */
/* OUTPUT                                                                                                                 */
/*                                                                                                                        */
/*  Table of Contents                                                                                                     */
/*                                                                                                                        */
/*  Females  ,............................    1                                                                           */
/*  Males    ,..............................  2                                                                           */
/*                                                                                                                        */
/*                                                                                                                        */
/* FEMALES                                                                                                                */
/*                                                                                                                        */
/*                                                                                                                        */
/*  NAME            AGE     HEIGHT     WEIGHT                                                                             */
/*                                                                                                                        */
/*  Alfred           14         69      112.5                                                                             */
/*  Henry            14       63.5      102.5                                                                             */
/*  James            12       57.3         83                                                                             */
/*  Jeffrey          13       62.5         84                                                                             */
/*  John             12         59       99.5                                                                             */
/*  Philip           16         72        150                                                                             */
/*  Robert           12       64.8        128                                                                             */
/*  Ronald           15         67        133                                                                             */
/*  Thomas           11       57.5         85                                                                             */
/*  William          15       66.5        112                                                                             */
/* +--------------+                                                                                                       */
/* | -- RETURN -- |                                                                                                       */
/* +--------------+                                                                                                       */
/*                                                                                                                        */
/*  NAME            AGE     HEIGHT     WEIGHT                                                                             */
/*                                                                                                                        */
/*  Alice            13       56.5         84                                                                             */
/*  Barbara          13       65.3         98                                                                             */
/*  Carol            14       62.8      102.5                                                                             */
/*  Jane             12       59.8       84.5                                                                             */
/*  Janet            15       62.5      112.5                                                                             */
/*  Joyce            11       51.3       50.5                                                                             */
/*  Judy             14       64.3         90                                                                             */
/*  Louise           12       56.3         77                                                                             */
/*  Mary             15       66.5        112                                                                             */
/* +--------------+                                                                                                       */
/* | -- RETURN -- |                                                                                                       */
/* +--------------+                                                                                                       */
/*                                                                                                                        */
/**************************************************************************************************************************/

/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/





