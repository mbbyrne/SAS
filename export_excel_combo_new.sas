

%macro excel_output(	
						filepath=,
						filename=,
						template=,
						listdelim=|,
						dsetlist=,
						tablist=,
						tabdesclist=,
						commentcols=dm_comments|comments,
						lib = work,
						toc=ON,
						border=ON,
						freeze=ON,
						filter=ON,
						autofit=ON,
						headerbold=ON,
						headerlabel=ON);

%let start_time =  %sysfunc(datetime(),datetime22.);

option missing="" nonotes;

ods listing close;
ods results off;

	%macro isBlank(param);
	%sysevalf(%superq(param)=,boolean)
	%mend isBlank; 


%let dset_count=%sysfunc(countw(&dsetlist.,"|"));


/*
insert code from coding macro to make unix or windows path from entered path
*/
options noxwait;
data shLabels;
	format TABNAME NAME LABEL $256.;
	set _null_;

		array charvars _character_;
			do over charvars;
				charvars = '';
			end;

run;
		
data toc;
	format TABNAME DESCRIPTION DATE RECORDS $200.;
	set _null_;

	label 	TABNAME = "Sheet Name"
			DESCRIPTION = "Output Description"
			DATE = "Date Created"
			RECORDS = "Number of Observations";

		array charvars _character_;
			do over charvars;
				charvars = '';
			end;
run;



%if %upcase(&toc.) = ON %then %do;

	proc contents noprint
				data = toc
				out = pc_toc;
			quit;

			proc sort
				data = pc_toc
				out = pc_toc_sorted;
				by VARNUM;
			quit;

			
			data shLabels_toc;
				set pc_toc_sorted(keep = NAME LABEL);
				TABNAME = "Summary";
			run;

data shLabels;
				set shLabels shLabels_toc;
			run;


%end;

	%do i = 1 %to &dset_count.;
	
	%let dset = %scan(&dsetlist.,&i.,'|');
	%let tab = %scan(&tablist.,&i.,'|');
	%let tabdesc = %scan(&tabdesclist.,&i.,'|');

    
			proc contents noprint
				data = &lib..&dset.
				out = pc;
			quit;

			proc sort
				data = pc
				out = pc_&i.;
				by VARNUM;
			quit;

			data shLabels_temp;
				set pc_&i.(keep = NAME LABEL);
				TABNAME = "&tab.";
			run;

        
proc sql noprint;
    select 
        NAME,
        NAME,
        LABEL
        into
        :varilist separated by "|",
        :varilist_&i. separated by ",",
        :varillist separated by "|"
        from pc_&i.;
quit;
   

        %let vari_count = %sysfunc(countw(&varilist.,"|"));
        %convert_types(dsetin=&dset.);



          data header_&i.();
            
                
            %do vari = 1 %to &vari_count.;

            %let varin=%scan(&varilist.,&vari.,"|");
           

            

%let varil=%quote(%scan(&varillist.,&vari.,"|"));
    
            %if &varil.= %then %let varil = &varin.;


            label &varin. = %quote("&varil.");
            &varin. =  "FIRSTROW";
            %end;

           run;



%put NEW CODE;
       proc sql noprint;
    select 
        "TEST"
        %do varnum=1 %to &vari_count.;
        %let current_var = %scan(&varilist.,&varnum.,"|");
        ,max(length(&current_var.))
        %end;
        
        into
         
        :testvar 
        %do varnum=1 %to &vari_count.;
        ,:&dset._&varnum. separated by ''
        %end;
        
       from &dset.;
quit;

%put &dset. is dset;

%put &&&dset._1 is the first dset_1;
%put &&&dset._2 is the second dset_2;

       
        

			data toc_temp(keep = TABNAME DESCRIPTION DATE RECORDS);
				set pc_&i.(keep = MEMNAME NOBS CRDATE);


				TABNAME = "&tab.";
				RECORDS = left(vvalue(NOBS));
				DATE =vvalue(CRDATE);
				DESCRIPTION = "&tabdesc.";

				if _n_ = 1 then output;

			run;


			data toc;
				set toc toc_temp;
			run;

		%end;
    




options validvarname=v7;

%let warn1 = WARN;
%let warn2 = ING;


%if %sysfunc(fileexist("&filepath.\&filename..xlsx")) %then %do;
%put File already exists - overwriting;
%sysexec %str(del "&filepath.\&filename..xlsx");
%end;



%*Output Initial formatted file;

%include "\\dub-filer-02\saseg_eu\sasdata\Test\excel\style\test.sas";

ods excel file="&filepath.\&filename..xlsx"
style=icon_dm/* sapphire STATDOC*/
options(
    contents="OFF"
    index="OFF"
    formulas="OFF"
    frozen_headers="ON"
    frozen_rowheaders="OFF"
    gridlines="ON"
    
    autofilter="ALL"
    sheet_interval="proc"
    sheet_name="Table of Contents"
    tab_color="white"

    /*ABSOLUTE_ROW_HEIGHT ="17"*/
);

ods escapechar='~';

proc print noobs label
    data = toc
    style(header)={just=c vjust=c background=white foreground=CX00B050};
        var _all_ / style(data)={tagattr="format:@"};
quit;
  


%do dset_num = 1 %to &dset_count.;

    %let dset = %scan(&dsetlist.,&dset_num.,'|');
    %let tab = %scan(&tablist.,&dset_num.,'|');


proc sql noprint;
    select NAME
    
    into
    :namelist separated by ' '
    from pc_&dset_num.;
    quit;

    


ods excel options(
sheet_name="&tab."
tab_color="CX00B050"
);

%put stage 1;



proc report
    data = header_&dset_num.(firstobs=1 obs=1) nowd  
    style (header) = [just=c vjust=c background=white foreground=CX00B050];
       
    
    column &namelist.;
    
/*
%do y = 1 %to %sysfunc(countw(&namelist.));

        %put dset is &dset. y is &y. Both is &&&dset._&y..;
        %let var = %scan(&namelist.,&y.);
        %if %sysevalf(&&&dset._&y.. < 25) %then %let &dset._&y. = 10;
        %put Greater than 25 &&&dset._&y.. for &dset.;
        %put stage 2;
        %let width = %sysevalf(&&&dset._&y.. * 10);

    %if %sysevalf(&&&dset._&y.. >= 30) %then %do;
        define &var./style=[width=&width. just=c tagattr='wrap:no'];
    %end;
    */
            
/*%end;*/


    %put stage 3;
/*    define &var. /ORDER noprint order=internal style=[font_size=10pt CELLWIDTH = 0&percent. ] "Patient" CENTER;*/
    /*
    %do y = 1 %to %sysfunc(countw(&namelist.));
    %let var = %scan(&namelist.,&y.," ");
        define &var./ style=[width = 200] ;
    %end;
    */
    quit;
/*

    proc print noobs label
        data = &dset._header(firstobs=1 obs=1)
        style(header)={just=c vjust=c background=white foreground=CX00B050};
        var _all_ / style(data)={tagattr="format:@"};
    quit;
*/

%end;


ods excel close;

/*%goto earlystop;*/

%put goto missed;

%*Insert data;

	libname xlfile excel 
				path= "&filepath.\&filename..xlsx" scan_text=NO HEADER=NO;


	%let dsetcount = %sysfunc(countw(&dsetlist.,'|'));
	
	%if %isblank(&tablist.) %then %let tablist = &dsetlist.;

	%else %do;

	%if %sysfunc(countw(&tablist.,'|'))^= &dsetcount. %then %do;
		%put &warn1.&warn2.: tablist %sysfunc(countw(&tablist.,'|')) and dsetlist &dsetcount. must have the same number of items - exiting macro;
		%goto exitmacro;
	%end;
	%end;
	

	%if %isblank(&tabdesclist.) %then %let tabdesclist = &tablist.;

	%else %do;
	
		%if %sysfunc(countw(&tabdesclist.,'|'))^= &dsetcount. %then %do;
			%put &warn1.&warn2.: tabdesclist %sysfunc(countw(&tabdesclist.,'|')) and dsetlist &dsetcount. must have the same number of items - exiting macro;
			%goto exitmacro;
		%end;

	%end;
			
	%do j = 1 %to &dset_count.;

	
	%let dset = %scan(&dsetlist.,&j.,'|');
	%let tab = %scan(&tablist.,&j.,'|');
	%let tabdesc = %scan(&tabdesclist.,&j.,'|');


/*%convert_types(dsetin=&dset.);*/
%insert(dsetin=&dset.,tabin=&tab.);



    %end;
	

libname xlfile clear;

%earlystop:


%let end_time = %sysfunc(datetime(),datetime22.);

option notes;

%put excel_output started at: &start_time.;
%put excel_output finished at: &end_time.;

%exitmacro:

%mend excel_output;

/*END*/




%macro convert_types(libin=work,dsetin=);

%let varlistc=;
%let varlistn=;

%let vararrn=;
%let vararrc=;

proc contents noprint
    data = &dsetin.
    out = pc_temp;
quit;

proc sql noprint;
    select 
    NAME,
    LABEL,
    TYPE
    into
    :namelist separated by '|', 
    :labellist separated by '|', 
    :typelist separated by '|'
    from pc_temp;
    quit;


%*Loop over input variables;
%do var = 1 %to %sysfunc(countw(&namelist.));
%let varcurrent	= %scan(&namelist.,&var.);
%let vartype = %scan(&typelist.,&var.);

%let varnum	= %scan(&namelist.,&var.)_num;
%let varchar	= %scan(&namelist.,&var.)_char;

	

	%if &vartype. =2 %then %do;
		%let vararrn= &vararrn. &varcurrent._num;
		%let varlistc= &varlistc. &varcurrent.;
	%end;
		
	%else %if &vartype. =1 %then %do;
		%let varlistn= &varlistn. &varcurrent.;
		%let vararrc= &vararrc. &varcurrent._char;
	%end;

%end;


data temp;
set &libin..&dsetin.(rename = (

%do varren = 1 %to %sysfunc(countw(&namelist.));
%let varcurrent	= %scan(&namelist.,&varren.);

&varcurrent. = &varcurrent._
%end;
));


run;

data &libin..&dsetin.(drop = 
%do varren = 1 %to %sysfunc(countw(&namelist.));
%let varcurrent	= %scan(&namelist.,&varren.);
&varcurrent._
%end;
);


set temp;




%do varren = 1 %to %sysfunc(countw(&namelist.));
%let varcurrent	= %scan(&namelist.,&varren.,"|");
%let labelcurrent = %scan(&labellist.,&varren.,"|");
%if %isblank(&labelcurrent.) %then %let labelcurrent=&varcurrent.;

label &varcurrent. = "&labelcurrent.";
    &varcurrent. = left(vvalue(&varcurrent._));

%end;
    
	
	
run;



%mend convert_types;

%macro split_sheet();



		/*Account for cases where output obs count exceeds 65535. Split these tables into multiple tabs*/

		%let splitrow=65534;
	
		proc sql noprint;
			select count(*) into :obs_count from &lib..&dset.;
		quit;

		%put obs count is &obs_count.;

		%if &obs_count GT &splitrow. %then 
%do;

			%put NOTE: obs count was high so splitting;
				proc sql noprint;
					select ceil(&obs_count./&splitrow.) into :loop_count from &lib..&dset.;
				quit;

				%put &loop_count. is loop count;

					data remainder;
				set &lib..&dset.;
			run;

			data remainder;
				set splittest;
			run;


			

			data remainder;
				set splittest;
			run; 

			%put WARNING: what is going on?;
%do k=1 %to &loop_count.;

				%put WARNING: processing k &k.;

				

				
			data new remainder;
				set remainder;
					if _n_ <= &splitrow. then output new;
					else output remainder; 
			run;

			
		data xlfile.&tab._&k.;
			set new;
		run;

 

%end;
%end;

%mend split_sheet;


%macro insert(dsetin=,tabin=);


/*
proc contents
    data = &dsetin.
    out = pc_temp
    ;
    quit;

   proc sort
				data = pc_temp
				out = pc_temp_sorted;
				by VARNUM;
			quit;
        */

proc sql noprint;
select 
    NAME
    into
    :namelist separated by " "
    from pc_&j.;
quit;

%put namelist for pc_&j. is &namelist.;

proc contents
    data = xlfile."&tabin.$"n
    out = pc_xl
    ;
    quit;
/*
proc sql noprint;
select 
    NAME
    into
    :xlnamelist separated by " "
    from pc_xl;
quit;

%put xl namelist is &xlnamelist.;
%put namelist is &namelist.;
  */  
%let varcount=%sysfunc(countw(&namelist.," "));

%put varcount is &varcount.;

	data append_&j.;
        set &dsetin.
    (rename = 
(
%do xlnum = 1 %to &varcount.;
%let var = %scan(&namelist.,&xlnum.," ");
&var. = F&xlnum.
%end;
));

run;


/*

proc sql;
%do xlnum = 1 %to &varcount.;
        %let var = %sysfunc(scan(&namelist.,&xlnum.," "));
        %let xlvar = %sysfunc(scan(&xlnamelist.,&xlnum.," "));

update xlfile."&tabin.$"n
       
        %put &var. &xlvar.;
           set &xlvar. = (select &var. from &dsetin.(firstobs=1 obs=1));
        %end;
  */     

		
/*

        %do xlnum = 1 %to &varcount.;
        %let var = %sysfunc(scan(&namelist.,&xlnum.," "));
        %let xlvar = F&xlnum.;
%put xlvar is &xlvar.;
update xlfile."&tabin.$"n
set &xlvar. = (select &var. from &dsetin.(firstobs=2 obs=2));
    %end;
*/
/*
update xlfile."&tabin.$"n
       
             set &xlvar.=
         case when &xlvar.="X" then "missing_not"
         else &xlvar.
end;
  */       


/*           */
        
/*
data test_update;
    set dset1;
    if _n_ = 1 then output;
    run;

    proc sql;
        update test_update
        set CONT = 
        case when not missing(CONT) AND monotonic()>1 then 1
        else CONT=CONT
        end;

       update test_update
        set ID = 
        case when ID=125 then 2
        else 0
        end;

       update test_update
        set SEGMENT = 
        case when SEGMENT=125 then 0
        else 3
        end;

        quit;
*/


data main_&j.;
    set &dsetin.;
    if  _n_ > 1 then output;
        
run;


proc sql;


  %do xlnum = 1 %to &varcount.;
        %let var = %scan(&namelist.,&xlnum.," ");
        %let xlvar = F&xlnum.;

     

update xlfile."&tabin.$"n
       
             set &xlvar.=
/*         case when substr(&xlvar.,1,3)="AUG" AND substr(&xlvar.,length(&xlvar.)-2,3)="TAG" then (select &xlvar. from append_&j.(firstobs=1 obs=1))*/


        case when &xlvar. = "FIRSTROW" then (select &xlvar. from append_&j.(firstobs=1 obs=1))
        else &xlvar.
end;
%end;

        insert into xlfile."&tabin.$"n

        select &&&varilist_&j..
        from main_&j.

        ;


        quit;


%mend insert;
