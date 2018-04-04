/*GET SSN OR HIC NUMS FROM PYTHON*/
/**/
/*%include "H:/temp_hic.sas";*/
/**/
/*rsubmit;*/
/*proc sql;*/
/*select distinct strip("'"||strip(hic)||"'") into :hics separated by ', '*/
/*from HIC_PROF_LISTB*/
/*WHERE HIC ^= '';*/
/*quit;*/
/*endrsubmit;*/
LIBNAME RWORK SLIBREF = WORK SERVER = SASP;




/*RSUBMIT;*/
/*PROC SQL;*/
/*CREATE TABLE ADDYS AS*/
/*	SELECT DISTINCT KEY, BENEID, BENE_SSN_NUM, */
/*BENE_SRNM_NAME, */
/*BENE_GVN_NAME,*/
/*BENE_MLG_CNTCT_ADR,*/
/*BENE_MLG_CNTCT_ADR2,*/
/*BENE_MLG_CNTCT_ADR3,*/
/*BENE_MLG_CNTCT_ADR4,*/
/*BENE_MLG_CNTCT_ADR5,*/
/*BENE_MLG_CNTCT_ADR6,*/
/*CITY,*/
/*STATE,*/
/*BENE_MLG_CNTCT_ZIP_CD,*/
/*BENE_RSDNC_SSA_STD_STATE_CD,*/
/*BENE_BIRTH_DT,*/
/*BENE_SEX_IDENT_CD*/
/*FROM EDB.Z5_edb_wb_2017_11*/
/*WHERE BENE_SSN_NUM IN (&SSNS.);*/
/*QUIT;*/
/*ENDRSUBMIT;*/
/**/
/**/
/**/
/*RSUBMIT;*/
/*PROC SQL;*/
/*CREATE TABLE ADDYS AS*/
/*	SELECT DISTINCT KEY, BENEID, BENE_SSN_NUM, */
/*BENE_SRNM_NAME, */
/*BENE_GVN_NAME,*/
/*BENE_MLG_CNTCT_ADR,*/
/*BENE_MLG_CNTCT_ADR2,*/
/*BENE_MLG_CNTCT_ADR3,*/
/*BENE_MLG_CNTCT_ADR4,*/
/*BENE_MLG_CNTCT_ADR5,*/
/*BENE_MLG_CNTCT_ADR6,*/
/*CITY,*/
/*STATE,*/
/*BENE_MLG_CNTCT_ZIP_CD,*/
/*BENE_RSDNC_SSA_STD_STATE_CD,*/
/*BENE_BIRTH_DT,*/
/*BENE_SEX_IDENT_CD*/
/*FROM EDB.Dme_edb_wb_2017_11*/
/*WHERE BENE_SSN_NUM IN (&SSNS.);*/
/*QUIT;*/
/*ENDRSUBMIT;*/
/**/
/**/

/*rsubmit;*/
/*proc sql;*/
/*select distinct strip("'"||strip(BENE_SSN_NUM)||"'") into :SSNS separated by ', '*/
/*from ADDYS*/
/*WHERE BENE_SSN_NUM ^='';*/
/*quit;*/
/*endrsubmit;*/

rsubmit;

ods output Members=Members;
proc datasets library=edb memtype=data;
run;
quit;

proc sort data= members;
by descending LastModified;
run;

data out;
set members;
by descending LastModified;
if _n_ < 3;
run;

proc sql noprint;
select distinct strip(name) into: var1-:var2
from out;
quit;
%put &var1.;

endrsubmit;

RSUBMIT;

PROC SQL;
CREATE TABLE backesj.PTB_MCARE_BENES AS
	SELECT DISTINCT KEY as HIC, BENEID as recip_id, BENE_SSN_NUM as recip_ssn, 
BENE_SRNM_NAME as RECIP_NAME_last, 
BENE_GVN_NAME as RECIP_NAME_FIRST,
BENE_MDL_NAME as RECIP_NAME_MIDDLE_INIT,
BENE_MLG_CNTCT_ADR as recip_address_line_1,
BENE_MLG_CNTCT_ADR2 as recip_address_line_2,
/*BENE_MLG_CNTCT_ADR3,*/
/*BENE_MLG_CNTCT_ADR4,*/
/*BENE_MLG_CNTCT_ADR5,*/
/*BENE_MLG_CNTCT_ADR6, adr6 basically contains the city from the table*/
CITY as recip_city,
STATE as recip_state,
BENE_MLG_CNTCT_ZIP_CD as recip_zip_cd,
BENE_RSDNC_SSA_STD_STATE_CD as state_cd,
BENE_BIRTH_DT as recip_date_birth,
BENE_DEATH_DT as recip_date_death
/*'PTB' as SOURCE*/
FROM EDB.&var2.
WHERE KEY ^= '' 
OR BENEID ^= ''
OR BENE_SSN_NUM ^='';
QUIT;

ENDRSUBMIT;

RSUBMIT;

PROC SQL;
CREATE TABLE backesj.DME_MCARE_BENES AS
	SELECT DISTINCT KEY as HIC, BENEID as recip_id, BENE_SSN_NUM as recip_ssn, 
BENE_SRNM_NAME as RECIP_NAME_last, 
BENE_GVN_NAME as RECIP_NAME_FIRST,
BENE_MDL_NAME as RECIP_NAME_MIDDLE_INIT,
BENE_MLG_CNTCT_ADR as recip_address_line_1,
BENE_MLG_CNTCT_ADR2 as recip_address_line_2,
/*BENE_MLG_CNTCT_ADR3,*/
/*BENE_MLG_CNTCT_ADR4,*/
/*BENE_MLG_CNTCT_ADR5,*/
/*BENE_MLG_CNTCT_ADR6, adr6 basically contains the city from the table*/
CITY as recip_city,
STATE as recip_state,
BENE_MLG_CNTCT_ZIP_CD as recip_zip_cd,
BENE_RSDNC_SSA_STD_STATE_CD as state_cd,
BENE_BIRTH_DT as recip_date_birth,
BENE_DEATH_DT as recip_date_death
/*'DME' as SOURCE*/
FROM EDB.&var1.
WHERE KEY ^= '' 
OR BENEID ^= ''
OR BENE_SSN_NUM ^='';
QUIT;

ENDRSUBMIT;

LIBNAME RWORK SLIBREF = WORK SERVER = SASP;

Rsubmit;
/*filename currar "/saspext1/Z5/sasdata/common/current_ar_medi.txt";*/
%MACRO LIB_LOOP1;
%LET ST1=AL;
%LET ST2=AR;
%LET ST3=GA;
%LET ST4=LA;
%LET ST5=MS;
%LET ST6=NC;
%LET ST7=TN;
%LET ST8=WV;

%DO I=1 %TO 8;

%PUT LOOP1 &I.;

DATA CURR_&&ST&I.;
infile curr&&ST&I. truncover;
INPUT @1 current $CHAR04. ;
RUN;


********************** UNCOMMENT BELOW IF YOU WANT TO FORCE A REFRESH FOR A STATE ************************;

/*%LET ST=GA;*/
/*DATA CURR_&ST.;*/
/*SET CURR_&ST.;*/
/*CURRENT='0816';*/
/*RUN;*/

%END;

%MEND;

ENDRSUBMIT;


RSUBMIT;

%LIB_LOOP1;

ENDRSUBMIT;


RSUBMIT;

%MACRO LIB_LOOP2;

%LET ST1=AL;
%LET ST2=AR;
%LET ST3=GA;
%LET ST4=LA;
%LET ST5=MS;
%LET ST6=NC;
%LET ST7=TN;
%LET ST8=WV;

%DO I=1 %TO 8;

%PUT LOOP2 &I.;

proc sql noprint;
      select current into :curr
      from curr_&&ST&I..
      ;
quit;

	LIBNAME &&ST&I.._CLAIM "/saspext2/Z5/sasdata/MM/&&ST&I../&curr./claims" ;
	LIBNAME &&ST&I.._PROV "/saspext2/Z5/sasdata/MM/&&ST&I../&curr./provider" ;
	LIBNAME &&ST&I.._BENE "/saspext2/Z5/sasdata/MM/&&ST&I../&curr./eligibility" ;
		%IF &I.=6 %THEN %DO;
		LIBNAME &&ST&I.._BENE "/saspext2/Z5/sasdata/MM/&&ST&I../&curr./elig" ;
		%END;

%END;

%MEND;

ENDRSUBMIT;


RSUBMIT;

%LIB_LOOP2;

ENDRSUBMIT;


%MACRO LIB_LOOP3;

%LET ST1=AL;
%LET ST2=AR;
%LET ST3=GA;
%LET ST4=LA;
%LET ST5=MS;
%LET ST6=NC;
%LET ST7=TN;
%LET ST8=WV;

%DO I=1 %TO 8;

%PUT LOOP3 &I.;


proc sql noprint;
      select current into :curr
      from RWORK.curr_&&ST&I.
      ;
quit;

	LIBNAME &&ST&I.._CLAIM "/saspext2/Z5/sasdata/MM/&&ST&I../&curr./claims"  SERVER=SASP;
	LIBNAME &&ST&I.._PROV "/saspext2/Z5/sasdata/MM/&&ST&I../&curr./provider"  SERVER=SASP;
	LIBNAME &&ST&I.._BENE "/saspext2/Z5/sasdata/MM/&&ST&I../&curr./eligibility"  SERVER=SASP;
		%IF &I.=6 %THEN %DO;
		LIBNAME &&ST&I.._BENE "/saspext2/Z5/sasdata/MM/&&ST&I../&curr./elig" SERVER=SASP;
		%END;

%END;

%MEND;

%LIB_LOOP3;



RSUBMIT;

%MACRO LOOP;

%LET ST1 = AL;
%LET ST2 = GA;
%LET ST3 = LA;
%LET ST4 = WV;
%LET ST5 = NC;
%LET ST6 = AR;
%LET ST7 = MS;
%LET ST8 = TN;

%DO I = 1 %TO 4;

PROC SQL;
CREATE TABLE &&ST&I.._PMPM AS
SELECT DISTINCT 
'' AS HIC,
"&&st&i.. " as SOURCE,
recip_address_line_1,
recip_address_line_2,
recip_city,
recip_date_birth,
recip_date_death,
recip_id,
RECIP_NAME_FIRST,
RECIP_NAME_LAST,
RECIP_NAME_MIDDLE_INIT,
recip_ssn,
recip_state,
RECIP_TELEPHONE,
recip_zip_cd
FROM &&ST&I.._BENE.&&ST&I.._ELIGIBILITY_PMPM
QUIT;

%END;

%DO I = 5 %TO 5;

PROC SQL;
CREATE TABLE &&ST&I.._PMPM AS
SELECT DISTINCT 
'' AS HIC,
"&&st&i.. " as SOURCE,
recip_addr_line_1 as recip_address_line_1,
recip_addr_line_2 as recip_address_line_2,
recip_city,
recip_date_birth,
recip_date_death,
recip_id,
RECIP_NAME_FIRST,
RECIP_NAME_LAST,
RECIP_NAME_MIDDLE_INIT,
recip_ssn,
recip_state,
RECIP_TELEPHONE,
recip_zip_cd
FROM &&ST&I.._BENE.RECIP_CUMUL_2017_12;
QUIT;

%END;

%DO I = 6 %TO 6;

PROC SQL;
CREATE TABLE &&ST&I.._PMPM AS
SELECT DISTINCT 
'' AS HIC,
"&&st&i.. " as SOURCE,
recip_address_line_1,
recip_address_line_2,
recip_city,
recip_date_birth,
recip_date_death,
recip_id,
RECIP_NAME_FIRST,
RECIP_NAME_LAST,
RECIP_NAME_MIDDLE_INIT,
recip_ssn,
recip_state,
"" AS RECIP_TELEPHONE,
recip_zip_cd
FROM &&ST&I.._BENE.&&ST&I.._ELIGIBILITY_PMPM;
QUIT;

%END;

%DO I = 7 %TO 8;

PROC SQL;
CREATE TABLE &&ST&I.._PMPM AS
SELECT DISTINCT 
'' AS HIC,
"&&st&i.. " as SOURCE,
recip_address_line_1,
recip_address_line_2,
recip_city,
recip_date_birth,
recip_date_death,
recip_id,
RECIP_NAME_FIRST,
RECIP_NAME_LAST,
RECIP_NAME_MIDDLE_INIT,
recip_ssn,
recip_state,
"" AS RECIP_TELEPHONE,
recip_zip_cd
FROM &&ST&I.._BENE.&&ST&I.._ELIGIBILITY_PMPM;
QUIT;

%END;


DATA BACKESJ.MCAID_BENES;
length 
RECIP_NAME_LAST $35.
RECIP_NAME_FIRST $30. 
RECIP_NAME_MIDDLE_INIT $15.
recip_address_line_1 recip_address_line_2 $60.
recip_city $30.
recip_id $15.
RECIP_ZIP_CD $9.;
SET 
AL_PMPM
GA_PMPM
LA_PMPM
WV_PMPM
NC_PMPM
AR_PMPM
MS_PMPM
TN_PMPM
;
WHERE HIC ^= '' 
OR recip_id ^= ''
OR recip_ssn ^='';
DROP SOURCE;
RUN;


%MEND;

ENDRSUBMIT;

RSUBMIT;

%LOOP;

ENDRSUBMIT;


* COMBINE BENES;

RSUBMIT;

DATA ALL_BENES;
LENGTH
RECIP_NAME_LAST $35.
RECIP_NAME_FIRST $30. 
RECIP_NAME_MIDDLE_INIT $15.
recip_address_line_1 recip_address_line_2 $60.
recip_city $30.
recip_id $15.
RECIP_ZIP_CD $9.;
SET 
BACKESJ.PTB_MCARE_BENES (OBS=100000)
BACKESJ.DME_MCARE_BENES (OBS=100000)
BACKESJ.MCAID_BENES (OBS=100000);
IF FINDW( UPCASE(recip_address_line_1) , STRIP(UPCASE(RECIP_CITY)) ) > 0 AND
FINDW( UPCASE(recip_address_line_1) , STRIP(UPCASE(RECIP_STATE)) ) > 0 THEN recip_address_line_1 = '';
IF FINDW( UPCASE(recip_address_line_2) , STRIP(UPCASE(RECIP_CITY)) ) > 0 AND
FINDW( UPCASE(recip_address_line_2) , STRIP(UPCASE(RECIP_STATE)) ) > 0 THEN recip_address_line_2 = '';
IF FINDW( UPCASE(recip_address_line_1) , STRIP(UPCASE(RECIP_NAME_LAST)) ) > 0 THEN 
recip_address_line_2 = 
TRANWRD(recip_address_line_1, STRIP(UPCASE(RECIP_NAME_LAST)), '');
IF FINDW( UPCASE(recip_address_line_1) , STRIP(UPCASE(RECIP_NAME_FIRST)) ) > 0 THEN 
recip_address_line_2 = 
TRANWRD(recip_address_line_1, STRIP(UPCASE(RECIP_NAME_FIRST)), '');
IF FINDW( UPCASE(recip_address_line_2) , STRIP(UPCASE(RECIP_NAME_LAST)) ) > 0 THEN recip_address_line_2 = 
TRANWRD(recip_address_line_2, STRIP(UPCASE(RECIP_NAME_LAST)), '');
IF FINDW( UPCASE(recip_address_line_2) , STRIP(UPCASE(RECIP_NAME_FIRST)) ) > 0 THEN recip_address_line_2 = 
TRANWRD(recip_address_line_2, STRIP(UPCASE(RECIP_NAME_FIRST)), '');

/*new = FIND(UPCASE(COMPRESS(recip_address_line_1," #")), 'POBOX');*/
/*IF FINDW( UPCASE(COMPRESS('PO BOX'," ")) , 'POBOX') > 0 THEN FLAG4=1;*/
	IF 
	FIND( UPCASE(COMPRESS(recip_address_line_1,"., #")) , 'POBOX') > 0 OR 
	FIND( UPCASE(recip_address_line_1) , 'PMB') > 0 OR
	FIND( UPCASE(COMPRESS(recip_address_line_1,"., ")) , 'OFFICEBOX') > 0 OR 
	FIND( UPCASE(COMPRESS(recip_address_line_1,"., ")) , 'POSTOFFICE') > 0 OR 
	FIND( UPCASE(COMPRESS(recip_address_line_1,"., ")) , 'POSTBOX') > 0 OR 
	FIND( UPCASE(COMPRESS(recip_address_line_1,"., ")) , 'BX') > 0 OR 
	FIND( UPCASE(COMPRESS(recip_address_line_1,"., ")) , 'BOX') > 0 OR 
	FIND( UPCASE(COMPRESS(recip_address_line_1,"., ")) , 'C/O') > 0 OR 
	FIND( UPCASE(recip_address_line_1) , 'LOT ') > 0 OR
	recip_address_line_1 = '' or
	FIND( UPCASE(recip_address_line_1) , 'UNIT ') > 0 OR
	FIND( UPCASE(recip_address_line_1) , 'SUITE') > 0 OR
	FIND( UPCASE(recip_address_line_1) , 'STE ') > 0 OR
	FIND( UPCASE(recip_address_line_1) , 'BUILDING') > 0 OR
	FIND( UPCASE(recip_address_line_1) , 'BLDG') > 0 OR
	( FIND( UPCASE(recip_address_line_1) , 'APT') > 0 AND LENGTH(recip_address_line_1) le LENGTH(recip_address_line_2)+3 ) or
	( FIND( UPCASE(recip_address_line_1) , 'RM') > 0 AND LENGTH(recip_address_line_1) le LENGTH(recip_address_line_2)+3 ) or
	( FIND( UPCASE(recip_address_line_1) , 'ROOM') > 0 AND LENGTH(recip_address_line_1) le LENGTH(recip_address_line_2)+3 )
	THEN DO;
	FLAG = 1;
	TEMP = recip_address_line_1;
	recip_address_line_1 = recip_address_line_2;
	recip_address_line_2 = TEMP;
END;

RUN;

ENDRSUBMIT;


RSUBMIT;

PROC SQL;
CREATE TABLE BACKESJ.ALL_BENES_DIST AS
SELECT DISTINCT *
FROM ALL_BENES
ORDER BY HIC, RECIP_SSN;
QUIT;

ENDRSUBMIT;


RSUBMIT;

PROC SQL;
CREATE TABLE backesj.DEAD_BENES AS
SELECT DISTINCT 
HIC, 
RECIP_ID, 
RECIP_SSN, 
MIN(recip_date_death) AS MIN_DEATH, 
MAX(recip_date_death) AS MAX_DEATH,
'Dead' as status 
FROM backesj.ALL_BENES_DIST
WHERE recip_date_death ^= .
GROUP BY 
HIC, 
RECIP_ID, 
RECIP_SSN;
QUIT;

ENDRSUBMIT;




RSUBMIT;

PROC SQL;
CREATE TABLE backesj.LIVE_BENES AS
SELECT DISTINCT *
FROM backesj.ALL_BENES_DIST
WHERE HIC NOT IN (SELECT DISTINCT HIC FROM BACKESJ.DEAD_BENES)
AND RECIP_ID NOT IN (SELECT DISTINCT RECIP_ID FROM BACKESJ.DEAD_BENES)
AND RECIP_SSN NOT IN (SELECT DISTINCT RECIP_SSN FROM BACKESJ.DEAD_BENES);
QUIT;

ENDRSUBMIT;

/*This script uses fuzzy matching of addresses and is based of the paper found below*/

/*http://www.scsug.org/wp-content/uploads/2013/11/The-Complexities-of-an-Address-Barry-Mullins.pdf*/

/*It takes a list of address and strips them down so they can be matched on their core parts.
Once they are match on one another they are given a score to judge their likeness and thier duplicates deleted.
We can then group the similar address together and create a summary to see how many unique benes 
are being billed from that address on the same day.*/

/****************************************************************************************************************/
/****************************************************************************************************************/


rsubmit;
proc sql;
create table BACKESJ.LIVE_BENES2 as
select distinct *, 
				cat(RECIP_NAME_FIRST,'',RECIP_NAME_LAST) as bene_name, 
				cat(recip_address_line_1,'', recip_address_line_2) as bene_address,  			/*For some its easier to just use the first address column as the second is sometimes the city or provider name*/
				substr(recip_zip_cd,1,5) as recip_zip_cd_5,
				case when length(recip_zip_cd) > 5 then recip_zip_cd end as recip_zip_cd_9
				
from BACKESJ.LIVE_BENES;
quit;
endrsubmit;

rsubmit;

data temp;
set backesj.live_benes2 (obs = 100000);
run;
endrsubmit;


/*Creating Macro to standardize the address variables*/
rsubmit;
%MACRO ADDRESS (VAR=,VAR2=); 
&VAR = TRANWRD(&VAR,"NORTH ","N "); 
&VAR = TRANWRD(&VAR,"NORTHWEST ","NW "); 
&VAR = TRANWRD(&VAR,"NORTHEAST ","NE ");
&VAR = TRANWRD(&VAR,"SOUTH ","S "); 
&VAR = TRANWRD(&VAR,"SOUTHWEST ","SW "); 
&VAR = TRANWRD(&VAR,"SOUTHEAST ","SE "); 
&VAR = TRANWRD(&VAR,"EAST ","E "); 
&VAR = TRANWRD(&VAR,"WEST ","W ");

If _n_=1 then do; 
keys= prxparse("s/\sST |\sSTREET |\sAVE |\sAV |\sAVENUE |\sDR |\sDRIVE |\sLN |\sLANE |\sRD |\sROAD 
		|\sPKWY |\sPARKWAY |\sBLVD |\sBOULEVARD |\sPL |\sPLACE |\sPLAZA |\sCT |\sCRT |\sCOURT |\sCIR |\sCIRCLE / /I");
end; 
retain keys;
&VAR2 = prxchange(keys,-1,&VAR);

&VAR2 = LEFT(PRXCHANGE('s/ #|STE | APT | UNIT | SUITE | BUILDING | BLDG / ZTE /',-1,CAT(' ',&VAR2,' '))); 
&VAR2 = LEFT(PRXCHANGE('s/ FRWY | FREEWAY / FWY /',-1,CAT(' ',&VAR2,' '))); 
&VAR2 = LEFT(PRXCHANGE('s/ EXPRWY | EXPRESSWAY | EXPWY / EXPY /',-1,CAT(' ',&VAR2,' ')));
&VAR2 = LEFT(PRXCHANGE('s/ HIWAY | HIGHWAY / HWY /',-1,CAT(' ',&VAR2,' ')));
&VAR2 = LEFT(PRXCHANGE('s/ RM | ROOM | N ROOM | S ROOM | E ROOM | W ROOM / ZRM /',-1,CAT(' ',&VAR2,' '))); 
&VAR2 = LEFT(PRXCHANGE('s/ P O BOX | PO BOX | POBOX | P0 BOX | P OBOX | Po BOX | Po Box | P O Box | PO BOS | BOX | POB | PMB | P O 
		DRAWER | PO DRAWER | POST OFFICE DRAWER | DRAWER | PBS BOX / ZPB /',-1,CAT(' ',&VAR2,' '))); 
&VAR2 = TRANWRD(&VAR2,"ROUTE ","RTE ");
%MEND ADDRESS;
endrsubmit;

rsubmit;
%MACRO BREAKUP_ADD (PATTERN=,VAR=,NUM=,NEWVAR=);
IF _N_=1 THEN DO; 
RETAIN ExpID&NUM; PATTERN="/&PATTERN/I"; 
ExpID&NUM=PRXPARSE(PATTERN); 
END;
CALL PRXSUBSTR(ExpID&NUM, &VAR, POSITION&NUM);

IF POSITION&NUM = 1 THEN DO; 
MATCH = SUBSTR(&VAR,POSITION&NUM); 
&NEWVAR=MATCH; END;
IF INDEX(&VAR,"&PATTERN") THEN &NEWVAR=SUBSTR(&VAR,POSITION&NUM); 
IF INDEX(&VAR,"&PATTERN") THEN SUBSTR(&VAR,POSITION&NUM)="";
%MEND BREAKUP_ADD;
endrsubmit;

/*Changing APT and UNIT for first list*/

		rsubmit;
		data random11;
		set backesj.live_benes2;
		ad=findw(upcase(bene_address), "APT");
		ad2=substr(bene_address, ad,ad+6);
		apt_num=substr(ad2,1,7);
		_apt_removed = compbl(tranwrd(upcase(bene_address),trim(apt_num), '') );
		new_address = catx(' ', propcase(_apt_removed), apt_num);
		bene_address2=new_address;
/*		drop ad ad2 _apt_removed new_address;*/
		run;
		endrsubmit;
/*	rsubmit;*/
/*		proc sql;*/
/*		drop table backesj.live_benes2;*/
/*		quit;*/
/*		endrsubmit;*/
		rsubmit;
		data random12;
		set random11;
		ad=findw(upcase(bene_address2), "UNIT");
		ad2=substr(upcase(bene_address2), ad,ad+6);
		apt_num=upcase(substr(ad2,1,7));
		_apt_removed = compbl(tranwrd(upcase(bene_address2),trim(apt_num), '') );
		new_address = catx(' ', propcase(_apt_removed), apt_num);
		bene_address3=new_address;
		drop ad ad2 _apt_removed new_address;
		run;
		endrsubmit;
/*		rsubmit;*/
/*		proc sql;*/
/*		drop table random11;*/
/*		quit;*/
/*		endrsubmit;*/

		/*Running Macro created before to create the first list*/
			rsubmit;
			data list1;
			set random12;

			LIST1_NAME=COMPRESS(COMPBL(bene_name),".,*<>"); 
			STREET_LIST1=UPCASE(COMPRESS(COMPBL(bene_Address3),".,*<>")); 
			ZIP5=substr(RECIP_ZIP_CD,1,5);

/*			drop hic bene_name bene_address city bene_state bene_zip bene_death_dt bene_birth_dt bene_sex bene_address3;*/
/*DROP */
/*HIC RECIP_ID RECIP_SSN */
/*BENE_ADDRESS RECIP_CITY RECIP_STATE RECIP_ZIP_CD RECIP_DATE_DEATH*/
/*RECIP_DATE_BIRTH BENE_ADDRESS3*/
/*SOURCE BENE_NAME KEYS APT_NUM BENE_ADDRESS2 */
/*EXPID1*/
/*EXPID2*/
/*EXPID3*/
/*EXPID4*/
/*EXPID5*/
/*POSITION1*/
/*POSITION2*/
/*POSITION3*/
/*POSITION4*/
/*POSITION5*/
/**/
/*PATTERN; */
			%ADDRESS (VAR=STREET_LIST1,VAR2=STREET1);
			%BREAKUP_ADD (PATTERN=ZPB,VAR=STREET1,NUM=1,NEWVAR=PO_BOX_STREET); 
			%BREAKUP_ADD (PATTERN=ZTE,VAR=STREET1,NUM=2,NEWVAR=SUITE_STREET); 
			%BREAKUP_ADD (PATTERN=ZTE,VAR=PO_BOX_STREET,NUM=3,NEWVAR=SUITE_STREET); 
			%BREAKUP_ADD (PATTERN=ZRM,VAR=STREET1,NUM=4,NEWVAR=RM_STREET); 
			%BREAKUP_ADD (PATTERN=ZRM,VAR=SUITE_STREET,NUM=5,NEWVAR=RM_STREET);

			PO_BOX_STREET = STRIP(TRANWRD(PO_BOX_STREET,'ZPB','')); 
			SUITE_STREET = STRIP(TRANWRD(SUITE_STREET,'ZTE','')); 
			RM_STREET = STRIP(TRANWRD(RM_STREET,'ZRM',''));
			PO_BOX_STREET=COMPRESS(PO_BOX_STREET,"#"); 
			SUITE_STREET=COMPRESS(SUITE_STREET,"#");
			run;
			endrsubmit;
/*		rsubmit;*/
/*		proc sql;*/
/*		drop table random12;*/
/*		quit;*/
/*		endrsubmit;*/
			rsubmit;
			PROC SORT DATA=LIST1 OUT=LIST1_NODUP NODUP; 
			BY STREET1;
			RUN;
			endrsubmit;
					rsubmit;

		proc sql;
		drop table LIST1;
		quit;
		endrsubmit;

/*Creating the second list by slightly changing the variable names so we can match later*/


************* DO THE SAME FOR FAC LIST **************;

* IMPORT FAC LISTS ****;

		rsubmit;
data backesj.LIST1_NODUP;
set LIST1_NODUP;
run;
		endrsubmit;

/*RSUBMIT;*/
/*PROC FREQ DATA = BACKESJ.ALL2;*/
/*TABLES recip_address_line_2/ OUT=JAY;*/
/*QUIT;*/
/*ENDRSUBMIT;*/

rsubmit;
proc sql;
create table BACKESJ.ALL2 as
select distinct *, 
				cat(recip_address_line_1,'', recip_address_line_2) as bene_address,  			/*For some its easier to just use the first address column as the second is sometimes the city or provider name*/
				substr(recip_zip_cd,1,5) as recip_zip_cd_5,
				case when length(recip_zip_cd) > 5 then recip_zip_cd end as recip_zip_cd_9
				
from BACKESJ.ALL;
quit;
endrsubmit;


/*Changing APT and UNIT for first list*/

		rsubmit;
		data random11B;
		set BACKESJ.ALL2;
		ad=findw(upcase(bene_address), "APT");
		ad2=substr(bene_address, ad,ad+6);
		apt_num=substr(ad2,1,7);
		_apt_removed = compbl(tranwrd(upcase(bene_address),trim(apt_num), '') );
		new_address = catx(' ', propcase(_apt_removed), apt_num);
		bene_address2=new_address;
		drop ad ad2 _apt_removed new_address;
		run;
		endrsubmit;
	rsubmit;
		proc sql;
		drop table backesj.all2;
		quit;
		endrsubmit;
		rsubmit;
		data random12B;
		set random11B;
		ad=findw(upcase(bene_address2), "UNIT");
		ad2=substr(upcase(bene_address2), ad,ad+6);
		apt_num=upcase(substr(ad2,1,7));
		_apt_removed = compbl(tranwrd(upcase(bene_address2),trim(apt_num), '') );
		new_address = catx(' ', propcase(_apt_removed), apt_num);
		bene_address3=new_address;
		drop ad ad2 _apt_removed new_address;
		run;
		endrsubmit;

	rsubmit;
		proc sql;
		drop table random11B;
		quit;
		endrsubmit;

		/*Running Macro created before to create the first list*/
			rsubmit;
			data list1B;
			set random12B;

			LIST1_NAME=COMPRESS(COMPBL(bene_name),".,*<>"); 
			STREET_LIST1=UPCASE(COMPRESS(COMPBL(bene_Address3),".,*<>")); 
			ZIP5=substr(RECIP_ZIP_CD,1,5);

/*			drop hic bene_name bene_address city bene_state bene_zip bene_death_dt bene_birth_dt bene_sex bene_address3;*/
DROP 
/*HIC RECIP_ID RECIP_SSN */
BENE_ADDRESS RECIP_CITY RECIP_STATE RECIP_ZIP_CD RECIP_DATE_DEATH
BENE_ADDRESS3
SOURCE BENE_NAME KEYS APT_NUM BENE_ADDRESS2 
EXPID1
EXPID2
EXPID3
EXPID4
EXPID5
POSITION1
POSITION2
POSITION3
POSITION4
POSITION5

PATTERN; 
			%ADDRESS (VAR=STREET_LIST1,VAR2=STREET1);
			%BREAKUP_ADD (PATTERN=ZPB,VAR=STREET1,NUM=1,NEWVAR=PO_BOX_STREET); 
			%BREAKUP_ADD (PATTERN=ZTE,VAR=STREET1,NUM=2,NEWVAR=SUITE_STREET); 
			%BREAKUP_ADD (PATTERN=ZTE,VAR=PO_BOX_STREET,NUM=3,NEWVAR=SUITE_STREET); 
			%BREAKUP_ADD (PATTERN=ZRM,VAR=STREET1,NUM=4,NEWVAR=RM_STREET); 
			%BREAKUP_ADD (PATTERN=ZRM,VAR=SUITE_STREET,NUM=5,NEWVAR=RM_STREET);

			PO_BOX_STREET = STRIP(TRANWRD(PO_BOX_STREET,'ZPB','')); 
			SUITE_STREET = STRIP(TRANWRD(SUITE_STREET,'ZTE','')); 
			RM_STREET = STRIP(TRANWRD(RM_STREET,'ZRM',''));
			PO_BOX_STREET=COMPRESS(PO_BOX_STREET,"#"); 
			SUITE_STREET=COMPRESS(SUITE_STREET,"#");
			run;
			endrsubmit;
	rsubmit;
		proc sql;
		drop table random12B;
		quit;
		endrsubmit;
			rsubmit;
			PROC SORT DATA=LIST1B OUT=LIST1B_NODUP NODUP; 
			BY STREET1;
			RUN;
			endrsubmit;

rsubmit;

DATA LIST1_NODUP_FINAL;
SET LIST1_NODUP;
IF STREET1 ^= '' THEN ADDRESS = STREET1;
ELSE IF PO_BOX_STREET ^= '' THEN ADDRESS = PO_BOX_STREET;
ELSE IF SUITE_STREET ^= '' THEN ADDRESS = SUITE_STREET;
ELSE IF RM_STREET ^= '' THEN ADDRESS = RM_STREET;
DROP STREET_LIST1 MATCH STREET1 PO_BOX_STREET SUITE_STREET RM_STREET;
RUN;

endrsubmit;
			

rsubmit;

DATA LIST1B_NODUP_FINAL;
SET LIST1B_NODUP;
IF STREET1 ^= '' THEN ADDRESS = STREET1;
ELSE IF PO_BOX_STREET ^= '' THEN ADDRESS = PO_BOX_STREET;
ELSE IF SUITE_STREET ^= '' THEN ADDRESS = SUITE_STREET;
ELSE IF RM_STREET ^= '' THEN ADDRESS = RM_STREET;
DROP STREET_LIST1 MATCH STREET1 PO_BOX_STREET SUITE_STREET RM_STREET;
RUN;
endrsubmit;

* EXPORT FILES ;

rsubmit;
data backesj.list1b_nodup_final;
set LIST1b_NODUP_FINAL;
run;

endrsubmit;

proc export data=backesj.LIST1_NODUP_FINAL
outfile="F:/backesj/fullBeneList.csv"
dbms =csv ;
run;

proc export data=rwork.LIST1B_NODUP_FINAL
outfile="F:/backesj/facAddressList.xlsx"
dbms =xlsx replace label;
sheet="Sheet1";
run;




























	rsubmit;
		proc sql;
		drop table LIST1B;
		quit;
		endrsubmit;


/*Creating the second list by slightly changing the variable names so we can match later*/






			rsubmit;
proc sql;
create table list2_nodup as
select provider,
		LIST1_NAME as list2_name ,
		ZIP5 as zip5_2,
		STREET1 as street2,
		PO_BOX_STREET as po_box_street_2,
		SUITE_STREET as suite_street_2,
		RM_STREET as rm_street_2,
		thru_dt as thru_dt2
from list1_nodup
;
quit;
endrsubmit;
/*Matching the two lists together and calculating a score*/
/*The score is determined how well the two main street address components match.
The lower the score the better the match. Uses the Compged function.*/
rsubmit;
Proc sql noprint ; 
create table STANDARD_STREET_MATCH as
select  A.LIST1_NAME ,
		A.ZIP5 ,
		A.STREET1 ,
		A.PO_BOX_STREET ,
		A.SUITE_STREET ,
		A.RM_STREET ,
		a.thru_dt as thru_dt1,
		B.LIST2_NAME ,
		B.ZIP5_2 ,
		B.STREET2,
		B.PO_BOX_STREET_2 , 
		B.SUITE_STREET_2 ,
		B.RM_STREET_2 , 
		b.thru_dt2,
		b.provider,
		compged(A.STREET1,B.STREET2,400,'I') AS SCORE
From LIST1_NODUP AS A, LIST2_NODUP AS B
WHERE A.ZIP5=B.ZIP5_2 AND A.STREET1 NE "" AND CALCULATED SCORE lt 350
order by SCORE, a.street1; 
quit ;
endrsubmit;


/*Cleans data so benes don't match back to each other*/
rsubmit;
data clean;
set Standard_street_match;
if (List1_name=list2_name) then delete;
run;
proc sort data=clean;
by thru_dt1 street1 list1_name;
run;
endrsubmit;

/* This makes sure street names are exact matches, Sometimes gets rid of the typos when they 
should be included. Only use if you want 100% exact matches. Recommend not using.*/

/*
rsubmit;
data clean2;
set Standard_street_match;
if (List1_name=list2_name) then delete;
where street1=street2;
run;

proc sort data=clean2;
by thru_dt1 street1 list1_name ;
run;
endrsubmit;
*/

rsubmit;
proc sql;
create table counts as
select distinct thru_dt1, street1, count(distinct list2_name) as count , provider
from clean /*clean2*/
group by thru_dt1, street1, provider
order by count desc
 ;
quit;
endrsubmit;

rsubmit;
proc sql;
create table count_ge_3 as
select *, case when substr(provider,1,2) = '11' then 'GA'
		 when substr(provider,1,2) in ('19','71') then 'LA'
		 when substr(provider,1,2) = '04' then 'AR'
		 when substr(provider,1,2) = '25' then 'MS'
		 when substr(provider,1,2) = '34' then 'NC'
		 when substr(provider,1,2) = '42' then 'SC'
		 when substr(provider,1,2) = '44' then 'TN'
		 when substr(provider,1,2) = '49' then 'VA'
		 when substr(provider,1,2) = '51' then 'WV'
		 when substr(provider,1,2) = '01' then 'AL'
		else 'Not Entered' end as State		
from counts
where count ge 3;
endrsubmit;

/*Building provider tables*/

rsubmit;
%macro counts;
	%do i = 3 %to 8;
			proc sql;
			create table count_&i as
			select provider, case when substr(provider,1,2) = '11' then 'GA'
					 when substr(provider,1,2) in ('19','71') then 'LA'
					 when substr(provider,1,2) = '25' then 'MS'
					 when substr(provider,1,2) = '04' then 'AR'
					 when substr(provider,1,2) = '34' then 'NC'
					 when substr(provider,1,2) = '42' then 'SC'
					 when substr(provider,1,2) = '44' then 'TN'
					 when substr(provider,1,2) = '49' then 'VA'
					 when substr(provider,1,2) = '51' then 'WV'
					 when substr(provider,1,2) = '01' then 'AL'
					else 'Not Entered' end as State	,
					count(thru_dt1) as num_days&i label="Number of days with count=&i"	
			from counts
			where count = &i
			group by provider, state;
			quit;
	%end;
%mend counts;
endrsubmit;

rsubmit;
%counts;
endrsubmit;



rsubmit;
proc sql;
create table count_9 as
select provider, case when substr(provider,1,2) = '11' then 'GA'
             when substr(provider,1,2) in ('19','71') then 'LA'
             when substr(provider,1,2) = '25' then 'MS'
             when substr(provider,1,2) = '34' then 'NC'
             when substr(provider,1,2) = '04' then 'AR'
             when substr(provider,1,2) = '42' then 'SC'
             when substr(provider,1,2) = '44' then 'TN'
             when substr(provider,1,2) = '49' then 'VA'
             when substr(provider,1,2) = '51' then 'WV'
             when substr(provider,1,2) = '01' then 'AL'
            else 'Not Entered' end as State     ,
            count(thru_dt1) as num_days9 label="Number of days with count ge 9"     
from counts
where count ge 9
group by provider, state;
quit;
endrsubmit;


rsubmit;
proc sql;
create table summary1 as
select *
from count_3 as a left join count_4 as b
on a.provider=b.provider and a.state=b.state ;
quit;
endrsubmit;

rsubmit;
%macro summary;
	%do i=1 %to 5;
		%let j = %eval(&i + 1);
		%let k = %eval(&i + 4);


		proc sql;
		create table summary&j as
		select *
		from summary&i as a left join count_&k as b
		on a.provider=b.provider and a.state=b.state ;
		quit;
	%end;
%mend summary;
endrsubmit;

rsubmit;
%summary;
endrsubmit;

rsubmit;
data cover;
Count_Description ='Count refers to the number of benes who were billed that live at a similar address for that day'; output;
run;
endrsubmit;


/**************EXPORT************/

proc export data=rwork.cover
outfile="&filepath"
dbms =xlsx replace label;
sheet="Info";
run;

proc export data=rwork.counts
outfile="&filepath"
dbms =xlsx replace label;
sheet="Counts all";
run;

proc export data=rwork.count_ge_3
outfile="&filepath"
dbms =xlsx replace label;
sheet="Counts ge 3";
run;

proc export data=rwork.Summary6
outfile="&filepath"
dbms =xlsx replace label;
sheet="Summary by provider";
run;
















RSUBMIT;

PROC SQL;
CREATE TABLE backesj.DEAD_BENES AS
SELECT DISTINCT 
HIC, 
RECIP_ID, 
RECIP_SSN, 
MIN(recip_date_death) AS MIN_DEATH, 
MAX(recip_date_death) AS MAX_DEATH,
'Dead' as status 
FROM backesj.MCARE_BENES
WHERE recip_date_death ^= .
GROUP BY 
HIC, 
RECIP_ID, 
RECIP_SSN;
QUIT;

ENDRSUBMIT;

RSUBMIT;

PROC SQL;
CREATE TABLE backesj.MCARE_BENES2 AS
SELECT DISTINCT *
FROM backesj.MCARE_BENES
WHERE HIC NOT IN (SELECT DISTINCT HIC FROM BACKESJ.DEAD_MCARE_BENES)
AND RECIP_ID NOT IN (SELECT DISTINCT RECIP_ID FROM BACKESJ.DEAD_MCARE_BENES)
AND RECIP_SSN NOT IN (SELECT DISTINCT RECIP_SSN FROM BACKESJ.DEAD_MCARE_BENES)
AND RECIP_SSN NOT IN (SELECT DISTINCT RECIP_SSN FROM 
QUIT;

ENDRSUBMIT;





RSUBMIT;

PROC SQL;
CREATE TABLE BACKESJ.DEAD_MCAID_BENES AS
SELECT * FROM 



RSUBMIT;
PROC SQL;
CREATE TABLE ALL AS
SELECT DISTINCT 
recip_address_line_1,
recip_address_line_2,
recip_city,
recip_date_birth,
recip_date_death,
recip_id,
RECIP_NAME_FIRST,
RECIP_NAME_LAST,
RECIP_NAME_MIDDLE_INIT,
recip_ssn,
recip_state,
"" AS RECIP_TELEPHONE,
recip_zip_cd
FROM AR_BENE.AR_ELIGIBILITY_PMPM
WHERE RECIP_SSN ^='';
QUIT;

ENDRSUBMIT;




RSUBMIT;
%LET ST = WV;
ENDRSUBMIT;
RSUBMIT;

PROC SQL;
CREATE TABLE &ST._PMPM AS
SELECT DISTINCT
LTC_PROV_BMS_ID,
LTC_PROV_NAME,
RECIP_ETHNICITY_CD_WV,
RECIP_MCARE_HIC,
RECIP_NAME_FIRST,
RECIP_NAME_LAST,
RECIP_NAME_MIDDLE_INIT,
RECIP_RACE_CD,
RECIP_RACE_CD_WV,
RECIP_SEX_CD_WV,
RECIP_TELEPHONE,
RECIP_ZIP_PLUS_4,
recip_address_line_1,
recip_address_line_2,
recip_city,
recip_date_birth,
recip_date_death,
recip_id,
recip_id_bms_1,
recip_id_bms_2,
recip_id_bms_3,
recip_id_bms_4,
recip_id_bms_5,
recip_id_bms_6,
recip_id_bms_7,
recip_sex_cd_zpic,
recip_ssn,
recip_state,
recip_zip_cd
FROM &ST._BENE.&ST._ELIGIBILITY_PMPM
WHERE recip_ssn = '096522301';
QUIT;

ENDRSUBMIT;

ODS HTML;
PROC CONTENTS DATA= TN_BENE.TN_ELIGIBILITY_PMPM;
RUN;
QUIT;



RSUBMIT;
%LET ST = TN;
ENDRSUBMIT;

RSUBMIT;

PROC SQL;
CREATE TABLE &ST._PMPM AS
SELECT DISTINCT
RECIP_ADDR_COUNTY_CD_TN,
RECIP_LIVING_ARRG_CD_TN,
RECIP_MCARE_HIC,
RECIP_NAME_FIRST,
RECIP_NAME_FULL,
RECIP_NAME_LAST,
RECIP_NAME_MIDDLE_INIT,
RECIP_RACE_CD_TN,
RECIP_RACE_CD_ZPIC,
RECIP_SEX_CD_TN,
RECIP_ZIP_PLUS_4,
ltc_prov_id,
mdcare_prem_payor,
mdcare_type_cde,
parta_ind,
recip_addr_county_cd_fips,
recip_address_line_1,
recip_address_line_2,
recip_city,
recip_date_birth,
recip_date_death,
recip_id,
recip_id_current,
recip_id_former,
recip_id_orig,
recip_sex_cd_zpic,
recip_ssn,
recip_ssn_former,
recip_state,
recip_zip_cd
FROM &ST._BENE.&ST._ELIGIBILITY_PMPM
WHERE recip_ssn = '096522301';
QUIT;

ENDRSUBMIT;


RSUBMIT;
%LET ST = AL;
ENDRSUBMIT;
%LET ST = AR;

ODS HTML;
PROC CONTENTS DATA= &ST._BENE.&ST._RECIPIENT;
RUN;
QUIT;


RSUBMIT;

PROC SQL;
CREATE TABLE &ST._PMPM AS
SELECT DISTINCT


RECIP_SSN

PROC SQL;
 SELECT *
 FROM DICTIONARY.COLUMNS
 WHERE UPCASE(LIBNAME)='AR_BENE';
QUIT;


RSUBMIT;
PROC SQL;
CREATE TABLE A AS
 SELECT DISTINCT LIBNAME, MEMNAME 
 FROM DICTIONARY.COLUMNS
 WHERE UPCASE(LIBNAME) CONTAINS 'BENE'
 AND 
	(
		UPCASE(NAME) CONTAINS 'TELE' OR
		UPCASE(NAME) CONTAINS 'PHONE'
		);
QUIT;
ENDRSUBMIT;

RSUBMIT;

PROC SQL;
CREATE TABLE TEST AS
SELECT DISTINCT KEY FROM EDB.Z5_edb_wb_2017_11
WHERE KEY CONTAINS '55555';
QUIT;

ENDRSUBMIT;


