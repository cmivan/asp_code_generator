<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'=============================================================================='
'------------------------<< ������� ��д ������
'------------------------<< ��ϵQQ            ��619835864
'------------------------------------------------------------------------------'
'------------------------->> ��վ��������������<<------------------- '
'------------------------------------------------------------------------------'
'=============================================================================='


'---->>ϵͳ����
DIM SYS_PAGE,WEB_FOLDER_NUM

      WEB_INFO            = "��վ����ϵͳ"                 'ϵͳ��Ϣ
      WEB_DATE            = REPLACE(DATE(),"-","")          '����ת��
      WEB_PASSWORD        = "1"                            '��¼����
	  WEB_LOGIN_MAX_TIMES = 3                             '���Ƶ�¼����
	  WEB_DATABASE_TYPE   = ".MDB"                         '���ݿ��׺
	  WEB_PAGE_NUM        = 6                              '���ɵ�ҳ������·����
	  WEB_FOLDER_NUM      = 5                              '���ļ��и���


  DIM WEB_INFO,WEB_DATE,WEB_LOGIN_MAX_TIMES,WEB_DATABASE_TYPE,WEB_PAGE_NUM,NUM


REDIM WEB_FILE(WEB_PAGE_NUM)                               '��ҳ��·��   ����
REDIM WEB_PAGE(WEB_PAGE_NUM)                               '��ҳ�����   ����
REDIM SYS_MEUN_LINK(WEB_PAGE_NUM)                          '��ҳ��˵�   ����
REDIM WEB_FOLDER(WEB_FOLDER_NUM)                           '���ļ��и��� ����
               


'=========================================================================
'-->> ϵͳ���ɱ�Ҫ��Ŀ¼  WEB_FOLDER()
'=========================================================================

      WEB_FOLDER(0)  = "WEB_" & SESSION("WEB_MAIN_FOLDER")         '��վ-��Ŀ¼
      WEB_FOLDER(1)  = WEB_FOLDER(0) & "\ADMIN"                    '��վ-����ԱĿ¼
      WEB_FOLDER(2)  = WEB_FOLDER(0) & "\DATA_"  & WEB_DATE        '��վ-���ݿ�Ŀ¼
      WEB_FOLDER(3)  = WEB_FOLDER(0) & "\STYPE"                    '��վ-��ʽĿ¼
      WEB_FOLDER(4)  = WEB_FOLDER(0) & "\INC"                      '��վ-����Ŀ¼
      WEB_FOLDER(5)  = WEB_FOLDER(0) & "\JS"                       '��վ-JSĿ¼



'=========================================================================
'-->> ϵͳ���ɱ�Ҫ���ļ������� WEB_FILE()
'=========================================================================

      TABLE_NAME=UCASE(REQUEST.QueryString("EDIT_TABLE"))
	  
      WEB_FILE(0)  = WEB_FOLDER(3) & "\ADMIN_CSS.CSS"                            '�ļ�-��ʽҳ��
	  WEB_FILE(1)  = WEB_FOLDER(1) & "\ADMIN_" & TABLE_NAME &"_EDIT.ASP"         '�ļ�-�༭ҳ��
	  WEB_FILE(2)  = WEB_FOLDER(1) & "\ADMIN_" & TABLE_NAME &"_LIST.ASP"         '�ļ�-�б�ҳ��
      WEB_FILE(3)  = WEB_FOLDER(1) & "\ADMIN_" & TABLE_NAME &"_DETAIL.ASP"       '�ļ�-��ϸҳ��
      WEB_FILE(4)  = WEB_FOLDER(1) & "\ADMIN_" & TABLE_NAME &"_SEARCH.ASP"       '�ļ�-����ҳ��
	  WEB_FILE(5)  = WEB_FOLDER(1) & "\CONN_"  & WEB_DATE  & ".ASP"              '�ļ�-��������
      WEB_FILE(6)  = WEB_FOLDER(1) & "\ADMIN_ALL_PAGINATION.ASP"                 '�ļ�-��ҳ����
	  



'=====================================================================================
'-->> ����ҳ��ʹ��SESSION��¼Ҫ���ɵ�ҳ��[ ģ�� ]
'-->> ���ҳ��
'-->> �༭ҳ��
'-->> �б�ҳ��
'-->> ����ҳ��
'-->> ��������
'-->> ��ҳ����
'=====================================================================================
   FOR NUM=0 TO WEB_PAGE_NUM
      IF SESSION("PAGE_"&NUM)="" THEN SESSION("PAGE_"&NUM)=REQUEST.FORM("PAGE_"&NUM)
   NEXT



	
'=====================================================================================	
'-->> ��վ��Ŀ¼ ���� -->>  SESSION("WEB_MAIN_FOLDER")
'=====================================================================================
   IF SESSION("WEB_MAIN_FOLDER")="" THEN
	  SESSION("WEB_MAIN_FOLDER")=REPLACE(NOW(),"-","")                       '��վ��Ŀ¼
	  SESSION("WEB_MAIN_FOLDER")=REPLACE(SESSION("WEB_MAIN_FOLDER"),":","")  '��վ��Ŀ¼
	  SESSION("WEB_MAIN_FOLDER")=REPLACE(SESSION("WEB_MAIN_FOLDER")," ","")  '��վ��Ŀ¼
   END IF




'=====================================================================================
'-->> ϵͳѡ�����ݡ�ģ��ҳ��Ĳ˵���
'===================================================================================== 
IF SESSION("SYS_MENU")="" THEN

       SYS_MEUN_LINK(0)="��ʽҳ��"
       SYS_MEUN_LINK(1)="���/�༭"
       SYS_MEUN_LINK(2)="�б�ҳ��"
       SYS_MEUN_LINK(3)="��ϸҳ��"
	   SYS_MEUN_LINK(4)="����ҳ��"
	   SYS_MEUN_LINK(5)="��������"
       SYS_MEUN_LINK(6)="��ҳ��Ŀ"
	   
   FOR I=0 TO WEB_PAGE_NUM
       SYS_MENU=SYS_MENU & "<DIV><A HREF=""#"" onMouseMove=""JavaSCRIPT:ChangeDIV('"
	   SYS_MENU=SYS_MENU & I & "','JKDIV_'," &WEB_PAGE_NUM& ")"" style=""cursor:hand;"">"  
	   SYS_MENU=SYS_MENU & SYS_MEUN_LINK(I)
	   SYS_MENU=SYS_MENU & "</A></DIV>"
   NEXT
   
   '----------->>  �˵�ѡ����
       SESSION("SYS_MENU")=SYS_MENU
END IF












' ====================================================================================
'|------------------------------------------------------------------------------------|
'                                      �Զ��庯����                                  
'|------------------------------------------------------------------------------------|
' ====================================================================================


'==>> �������ֵ�
' -> $TABLE_NAME$         ,TABLE_NAME
' -> $TABLE_KEY$          ,TABLE_KEY
' -> $DATABASE_CONN$      ,DATABASE_CONN
		
	
	
'==>> �༭ҳ��� �ַ� => ��վ����
' -> $����������$ ,FIELD_F
' -> $д��������$ ,FIELD_W
' -> $��ȡ������$ ,FIELD_R
' -> $��������$ ,FIELD_TD
' -> $�����ж�$   ,FIELD_FORM
' -> $���ж�$   ,FIELD_JS







'============================================================================
'-->> ����ҳ��ĺ���(�˳�ҳ��\����ѡ�����ݿ�ҳ��)
'============================================================================

SUB QUIT()
    DIY_PAGE=REQUEST.QUERYSTRING("DIY_PAGE")
     '�ж�ѡ���ҳ��
	SELECT CASE DIY_PAGE
	
	       CASE "OUT"    '�˳�ҳ��
                SESSION.Abandon() 	
               	RESPONSE.REDIRECT("INDEX.ASP")
               	RESPONSE.END()
				
		   CASE	"RESELECT"  '����ѡ�����ݿ�
		        FOR NUM=0 TO WEB_PAGE_NUM
				    SESSION("PAGE_1")=""
				NEXT
		   
		   		
				SESSION("PAGE_1")=""
				SESSION("PAGE_2")=""
				SESSION("PAGE_3")=""
				SESSION("PAGE_4")=""
				SESSION("PAGE_5")=""                    '���ҳ��ģ��
				
		        SESSION("SYS_PAGE")="SELECT_DB"
                SESSION("MDATA_PATH")=""
	
               	RESPONSE.REDIRECT("INDEX.ASP")
               	RESPONSE.END()
    END SELECT
END SUB
CALL QUIT()





'============================================================================
'-->> ��ȡ�ֶεĺ���  
'============================================================================

SUB GET_FIELD(TABLE,G_TABLE,WHERE)
    IF TABLE<>"" AND TABLE=G_TABLE THEN
       '----------------
SET FILED_CONN = SERVER.CREATEOBJECT("ADODB.RECORDSET")
    FILED_CONN_STR = "SELECT * FROM " & TABLE
    FILED_CONN.OPEN FILED_CONN_STR,CONNSTR,0,1
       '----------------   
	   
	FOR I=0 TO FILED_CONN.FIELDS.COUNT-1
	
	
	   IF WHERE="" THEN
    RESPONSE.WRITE("<DIV CLASS='FILED_DIV'><A HREF='#'>")
	RESPONSE.WRITE(FILED_CONN.FIELDS.ITEM(I).NAME)	
	RESPONSE.WRITE("</A></DIV>")
	   ELSE
	RESPONSE.WRITE("<option VALUE='" & FILED_CONN.FIELDS.ITEM(I).NAME & "'>")
	RESPONSE.WRITE(FILED_CONN.FIELDS.ITEM(I).NAME)	
	RESPONSE.WRITE("</option>")  
	   END IF
	   
	   
    NEXT
	
       '---------------- 
    FILED_CONN.CLOSE
SET FILED_CONN = NOTHING	
    END IF  
END SUB







'============================================================================
'-->> ���ض�ȡ���ݿ���ֶ�����
'============================================================================
SUB GET_TYPE(NUM)
    SELECT CASE NUM
	       CASE "3"
		        RESPONSE.WRITE("����")
		   CASE "135"
		        RESPONSE.WRITE("ʱ��")
		   CASE "202"
		        RESPONSE.WRITE("�ı�")
		   CASE "203"
		        RESPONSE.WRITE("ע��")
    END SELECT
END SUB







'============================================================================
'-->> �����ļ�/Ŀ¼ �ĺ���(GO)    SUB WEB_CREATE("�ļ���","�ļ�����","����")
'============================================================================

SUB WEB_CREATE(S_NAME,S_CENTENT,S_TYPE)
    SELECT CASE S_TYPE
	
	       CASE "FOLDER"                          '�ļ��в�����������
 		        SET FSO=CREATEOBJECT("SCRIPTING.FILESYSTEMOBJECT")
                    IF FSO.FOLDEREXISTS(SERVER.MAPPATH(S_NAME))=FALSE THEN  
                       FSO.CREATEFOLDER(SERVER.MAPPATH(S_NAME))
                    END IF                            
				SET FSO=NOTHING 
	   
		   CASE "FILE"                            '�ļ��������������ļ�
		        SET FSO=CREATEOBJECT("SCRIPTING.FILESYSTEMOBJECT")
		        SET SAVEFILE=FSO.OPENTEXTFILE(SERVER.MAPPATH(S_NAME),2,TRUE)
		            SAVEFILE.WRITELINE(S_CENTENT)
				SET SAVEFILE=NOTHING
				SET FSO=NOTHING			
    END SELECT
END SUB






'============================================================================
'-->> ������Ϣ���   MSGBOX ����
'-->> MSG(STR,NUM) ,STR="��ʾ������" , NUM=1 (��ʾ) 
'-->> NUM=2 (��ʾ������)��NUM=3 (��ʾ������)
'============================================================================
SUB MSG(STR,NUM)
    SELECT CASE NUM
	       CASE 1
		   RESPONSE.Write("<SCRIPT>alert('"&STR&"');</SCRIPT>")
		   CASE 2
		   RESPONSE.Write("<SCRIPT>alert('"&STR&"');history.go(-1);</SCRIPT>")
		   RESPONSE.End()
		   CASE 3
		   RESPONSE.Write("<SCRIPT>alert('"&STR&"');</SCRIPT>")
		   RESPONSE.End()
	END SELECT
END SUB






'============================================================================
'-->> ������Ϣ���  STRTOCODE ����
'-->> �ַ��滻�����ض����ַ��滻��Ҫ���ɵ���վ����
'============================================================================

FUNCTION STRTOCODE(MYSTR)
	'==>> ������
		 '�����ַ� "&lt;" "&gt;"
          SESSION(MYSTR)=REPLACE(SESSION(MYSTR),"&lt;","<")
	      SESSION(MYSTR)=REPLACE(SESSION(MYSTR),"&gt;",">")
		 
	     MYSTR=REPLACE(MYSTR,"$TABLE_NAME$",TABLE_NAME)
		 MYSTR=REPLACE(MYSTR,"$TABLE_KEY$",TABLE_KEY)
		 MYSTR=REPLACE(MYSTR,"$DATABASE_CONN$",DATABASE_CONN)
		
	
	'==>> �༭ҳ���
	     MYSTR=REPLACE(MYSTR,"$����������$",FIELD_F)        '�༭ҳ��
		 MYSTR=REPLACE(MYSTR,"$д��������$",FIELD_W)        '�༭ҳ��
		 MYSTR=REPLACE(MYSTR,"$��ȡ������$",FIELD_R)        '�༭ҳ��
	     MYSTR=REPLACE(MYSTR,"$��������$",FIELD_TD)       '���ɱ��
		 MYSTR=REPLACE(MYSTR,"$�����ж�$",FIELD_FORM)       '�����ж�
		 MYSTR=REPLACE(MYSTR,"$���ж�$",FIELD_JS)         '���ж�
		 
		 
    '==>> �б�ҳ���	
		 MYSTR=REPLACE(MYSTR,"$ADMIN_LIST_DB_1$",ADMIN_LIST_DB_1)
		 MYSTR=REPLACE(MYSTR,"$ADMIN_LIST_DB_2$",ADMIN_LIST_DB_2)
		 
    '==>> ��ϸҳ���
	     MYSTR=REPLACE(MYSTR,"$ADMIN_DETAIL_DB_1$",ADMIN_DETAIL_DB_1)
	
	
	'==>> ���ݿ�����ҳ���
		 DATA_CONN_PATH  ="..\DATA_"  & WEB_DATE & "\DATA_" & WEB_DATE & WEB_DATABASE_TYPE	
		 MYSTR=REPLACE(MYSTR,"$DATA_CONN_PATH$",DATA_CONN_PATH)  '���ж�
		 
END FUNCTION




'============================================================================
'-->> ���ر��   <TABLE> ��<tr> ��<td> ��<div>��  TOTABLE ����
'-->> ָ�����ַ��滻��Ҫ���ɵ���վ HTML ����
'============================================================================

FUNCTION TOTABLE(TABLESTR,T_TYPE)
         TABLESTR="<"&T_TYPE&">"&TABLESTR&"</"&T_TYPE&">"
END FUNCTION
%>














<%
'================================================================
'=======================��¼�ж�=================================
'================================================================
LOGIN_PASS=REQUEST.FORM("LOGIN_PASS")
IF SESSION("SYS_PAGE")="" AND LOGIN_PASS<>"" THEN
    IF LOGIN_PASS="" THEN
       CALL MSG("=> "&NOW()&"\n \n=> �������¼����...",2)
ELSEIF LOGIN_PASS=WEB_PASSWORD THEN
       SESSION("SYS_PAGE")="SELECT_DB"
  ELSE
       SESSION("LOGIN_TIMES")=INT(SESSION("LOGIN_TIMES"))+1
IF SESSION("LOGIN_TIMES")>WEB_LOGIN_MAX_TIMES THEN
CALL MSG("=> "&NOW()&"\n \n=> ���Ѷ�δ����¼...\n \n=> IP�ѳɹ���¼��",3)
END IF
       CALL MSG("=> "&NOW()&"\n \n=> ��¼�������...\n \n=> ��"&SESSION("LOGIN_TIMES")&"�δ����¼",2)
   END IF
END IF




'================================================================
'=======================�ж�ҳ��=================================
'================================================================
   DATA_PATH=REQUEST.FORM("DATA_PATH")                      'ѡ������ݿ�·��     
IF DATA_PATH<>"" AND SESSION("MDATA_PATH")="" THEN
   SESSION("MDATA_PATH2")=DATA_PATH
   SESSION("MDATA_PATH")=REPLACE(REQUEST.FORM("DATA_PATH"),SERVER.MAPPATH("/"),"")
   SESSION("SYS_PAGE")="SHOW_DB_LIST"
   
      IF SESSION("MDATA_PATH") =SESSION("MDATA_PATH2") THEN
         SESSION("MDATA_PATH3")=SESSION("MDATA_PATH2")
      ELSE
	     SESSION("MDATA_PATH3")=".." & SESSION("MDATA_PATH")
      END IF
 



'================================================================
'===========����������ļ���=====================================
'================================================================
      FOR NUM=0 TO WEB_FOLDER_NUM
	      CALL WEB_CREATE(WEB_FOLDER(NUM),"","FOLDER")
      NEXT
    


'=============================================================
END IF






'=================================================================================
'============================��ϵͳ�����ݿ�����===================================
'=================================================================================

DIM CONNS,CONNSTR,TIME1,TIME2,MDB
    TIME1=TIMER   
	
	IF SESSION("MDATA_PATH")=SESSION("MDATA_PATH2") THEN
       MDATA_PATH=SESSION("MDATA_PATH2")
	   ELSE
	   MDATA_PATH=SERVER.MAPPATH("/") & SESSION("MDATA_PATH")
	END IF
	
    ON ERROR RESUME NEXT
    CONNSTR="DRIVER=MICROSOFT ACCESS DRIVER (*.MDB);DBQ=" + MDATA_PATH
    '------------------------ 	
SET CONNS=SERVER.CREATEOBJECT("ADODB.CONNECTION") 
    CONNS.OPEN CONNSTR
	
%>





<%
'=======================���������ʽ�ĺ���============================
FUNCTION SHOW_TYPE(STR,INTYPE,F_NAME)        
       DIM  SHOW_INPUT
 SELECT CASE INTYPE
	    CASE "text","file","password","radio","checkbox","hidden"
SHOW_INPUT="<input NAME=" & F_NAME & " TYPE='" & INTYPE & "' STYLE='WIDTH:100%' VALUE='" & STR & "' />"
		CASE "textarea"
SHOW_INPUT="<textarea NAME=""" & F_NAME & """ STYLE='WIDTH:100%' />" & STR & "</textarea>"
 END SELECT
	SHOW_TYPE=SHOW_INPUT
END FUNCTION



'=======================����JAVASCRIPT�ĺ���============================
FUNCTION SHOW_JS(S_NAME,S_STR,LIMIT)
       DIM JS_BACK_INFO
  SELECT CASE LIMIT 
  
            CASE "LIMIT_1"                  '�ж��Ƿ�Ϊ��
			
JS_BACK_INFO = JS_BACK_INFO & "if(document.ADMIN_EDIT_FORM.F_"&S_NAME&".value=="""")"
JS_BACK_INFO = JS_BACK_INFO & chr(13) & "{"
JS_BACK_INFO = JS_BACK_INFO & "alert('����д "&S_STR&" ��');"
JS_BACK_INFO = JS_BACK_INFO & chr(13)
JS_BACK_INFO = JS_BACK_INFO & "document.ADMIN_EDIT_FORM.F_"&S_NAME&".focus();"
JS_BACK_INFO = JS_BACK_INFO & chr(13)
JS_BACK_INFO = JS_BACK_INFO & "return false;"
JS_BACK_INFO = JS_BACK_INFO & chr(13) & "}"
JS_BACK_INFO = JS_BACK_INFO & chr(13) 
				 
            CASE "LIMIT_2"
			     
            CASE "LIMIT_3"

JS_BACK_INFO = JS_BACK_INFO & "var strm = document.ADMIN_EDIT_FORM.F_"&S_NAME&".value"
JS_BACK_INFO = JS_BACK_INFO & chr(13)
	
JS_BACK_INFO = JS_BACK_INFO & "var regm = /^"
JS_BACK_INFO = JS_BACK_INFO & "[a-zA-Z0-9_-]"
JS_BACK_INFO = JS_BACK_INFO & "+@[a-zA-Z0-9_-]"
JS_BACK_INFO = JS_BACK_INFO & "+(\.[a-zA-Z0-9"
JS_BACK_INFO = JS_BACK_INFO & "_-]+)+$/;"
	
JS_BACK_INFO = JS_BACK_INFO & chr(13)
JS_BACK_INFO = JS_BACK_INFO & "if (!strm.match(regm) &"
JS_BACK_INFO = JS_BACK_INFO & "& strm!=""" & """)"
JS_BACK_INFO = JS_BACK_INFO & chr(13) & "{"
JS_BACK_INFO = JS_BACK_INFO & chr(13) & "alert('�����ַ��ʽ������зǷ��ַ�!\n���飡');"
	
JS_BACK_INFO = JS_BACK_INFO & chr(13)
JS_BACK_INFO = JS_BACK_INFO & "document.ADMIN_EDIT_FORM.F_"&S_NAME&".focus();"
JS_BACK_INFO = JS_BACK_INFO & "return false;" 
JS_BACK_INFO = JS_BACK_INFO & chr(13) & "}"
JS_BACK_INFO = JS_BACK_INFO & chr(13)
	
	  
            CASE "LIMIT_4"

JS_BACK_INFO = JS_BACK_INFO & "var reg=/^\d+$/g; "				
JS_BACK_INFO = JS_BACK_INFO & chr(13) & "if(reg.test(document.ADMIN_EDIT_FORM.F_"&S_NAME&".value)==false)"				
JS_BACK_INFO = JS_BACK_INFO & chr(13) & "{"
JS_BACK_INFO = JS_BACK_INFO & chr(13) & "alert('"&S_STR&" ����д�������ֺţ�');"
JS_BACK_INFO = JS_BACK_INFO & chr(13) & "document.ADMIN_EDIT_FORM.F_"&S_NAME&".focus();"
JS_BACK_INFO = JS_BACK_INFO & "return false;"
JS_BACK_INFO = JS_BACK_INFO & chr(13) & "}"				
JS_BACK_INFO = JS_BACK_INFO & chr(13)
	
            CASE "LIMIT_5"
				 
JS_BACK_INFO = JS_BACK_INFO & chr(13) & "if(isDate(document.ADMIN_EDIT_FORM.F_"&S_NAME&")==false)"			 
JS_BACK_INFO = JS_BACK_INFO & chr(13) & "{"
JS_BACK_INFO = JS_BACK_INFO & chr(13) & "alert('"&S_STR&" ����дʱ�䣡');"
JS_BACK_INFO = JS_BACK_INFO & chr(13) & "document.ADMIN_EDIT_FORM.F_"&S_NAME&".focus();"
JS_BACK_INFO = JS_BACK_INFO & chr(13) &  "return false;"
JS_BACK_INFO = JS_BACK_INFO & chr(13) & "}"	
JS_BACK_INFO = JS_BACK_INFO & chr(13)
	
  END SELECT 
    SHOW_JS = JS_BACK_INFO
END FUNCTION




'---------�ж��ֶε����ͣ�����----->>
FUNCTION SHOW_INT(STR,INTYPE)
       DIM BACK_INFO
  SELECT CASE INTYPE 
         CASE "3" 
		       BACK_INFO=" INT("&STR&")"
		 CASE "202","203"
               BACK_INFO="CSTR("&STR&")"
		 CASE "135"
		 	   BACK_INFO=STR
  END SELECT
              SHOW_INT=BACK_INFO
END FUNCTION

%>





<%
' =================================================================================
'|=================================================================================|
'|==================================�����ļ�=======================================|
'|=================================================================================|
' =================================================================================
'FIELD_F  - $����������$ 
'FIELD_W  - $д��������$
'FIELD_R  - $��ȡ������$
'FIELD_TD - $��������$
'FIELD_FORM - $�����ж�$
'FIELD_JS   - $�����ж�$



 DATABASE_CONN="CONN_" & WEB_DATE & ".ASP"                            '���ݿ�����·��
 KEY_NUM=0

 IF TABLE_NAME<>"" AND REQUEST.Form("DB_FORM")="YES" THEN
	IF REQUEST.Form("DB_MAX")<>"" AND REQUEST.Form("DB_MAX")>0 THEN
	   
	   FOR G_NUM=0 TO REQUEST.Form("DB_MAX")
	   
	   FIELD_1=REQUEST.Form("FIELD_" & G_NUM)                    '�ֶ�����
	   FIELD_2=REQUEST.Form("FIELD_SHOW_" & G_NUM)               '��ʾ����
	   FIELD_3=REQUEST.Form("FIELD_TYPE_" & G_NUM)               '�ֶ�����
	   FIELD_4=REQUEST.Form("FIELD_MAX_" & G_NUM)                '��󳤶�
	   FIELD_5=REQUEST.Form("FIELD_INPUT_TYPE_" & G_NUM)         '���������
	   FIELD_6=REQUEST.Form("FIELD_INPUT_LIMIT_" & G_NUM)        '�ֶ�����
	   FIELD_7=REQUEST.Form("FIELD_INPUT_EMPEY_TYPE_" & G_NUM)   '�Ƿ����Ϊ��
	   FIELD_8=REQUEST.Form("FIELD_INPUT_CHECK_TYPE_" & G_NUM)   '�Ƿ�JS���

		   

'---------����Ǹ�������----->>
    IF FIELD_7="yes" THEN
       FIELD_7=FIELD_7
	   
ELSEIF FIELD_7="no" THEN
       FIELD_7=FIELD_7
	   
ELSEIF FIELD_7<>"" THEN
	   TABLE_KEY=FIELD_7
	   KEY_NUM=KEY_NUM+1
    END IF

 
	       
		   
IF FIELD_1<>"" AND TABLE_KEY<>FIELD_7 THEN            

'-�༭ҳ��(1)-$����������$---------------------------------  
FIELD_F=FIELD_F & "F_" & FIELD_1 & "=" & SHOW_INT("REQUEST.Form(""F_" & FIELD_1 & """)",FIELD_3)
FIELD_F=FIELD_F & chr(13)
		   
		   
'-�༭ҳ��(2)-$д��������$--------------------------------- 
FIELD_W=FIELD_W & "EDIT_CONN(""" & FIELD_1 & """)=F_" & FIELD_1
FIELD_W=FIELD_W & chr(13)
			 
		
		
'-�༭ҳ��(3)-$��ȡ������$---------------------------------   
FIELD_R=FIELD_R & "F_" & FIELD_1 & "=READ_CONN(""" & FIELD_1 & """)"
FIELD_R=FIELD_R & chr(13)
		   
		   
'-�༭ҳ��(4)-$��������$---------------------------------  
FIELD_TD=FIELD_TD & "<TR><TD align='right'>"
FIELD_TD=FIELD_TD & chr(13)
FIELD_TD=FIELD_TD & REQUEST.Form("FIELD_SHOW_" & G_NUM) & "��</TD>" & chr(13) &"<TD>"
FIELD_TD=FIELD_TD & chr(13)
			  
INTYPE  = FIELD_5                 '���������
F_NAME  = "F_" & FIELD_1          '�����NAME
				
FIELD_TD=FIELD_TD & SHOW_TYPE("<%=" & "F_" & FIELD_1 & replace("%$>","$",""),INTYPE,F_NAME)
FIELD_TD=FIELD_TD & chr(13)
FIELD_TD=FIELD_TD & "</TD></TR>"

		
'-�༭ҳ��(5)-$�����ж�$--------------------------------- 
FIELD_FORM = FIELD_FORM & chr(13)
FIELD_FORM = FIELD_FORM & "CALL FORM_CHECK(""" &FIELD_1& """,F_" &FIELD_2& ",""" &FIELD_6& """)"
	
	
'-�༭ҳ��(6)-$���ж�$--------------------------------- 	
IF FIELD_8<>"" THEN
   FIELD_JS=FIELD_JS & SHOW_JS(FIELD_1,FIELD_2,FIELD_6)
   FIELD_JS=FIELD_JS & SHOW_JS(FIELD_1,FIELD_2,FIELD_7)	
END IF
		




'--------------------------------------------------------------
'-�б�ҳ��-$ADMIN_LIST_DB_1$---------------------------------
'--------------------------------------------------------------
TEMP_FIELD_1=FIELD_1

CALL TOTABLE(FIELD_2,"TD")
     ADMIN_LIST_DB_1=ADMIN_LIST_DB_1&FIELD_2
	 
	 FIELD_1="<"&"%=LIST_CONN(""" & FIELD_1 & """)%"&">"
CALL TOTABLE(FIELD_1,"TD")
     ADMIN_LIST_DB_2=ADMIN_LIST_DB_2&FIELD_1


'--------------------------------------------------------------
'-��ϸҳ��-$ADMIN_DETAIL_DB_1$---------------------------------
'--------------------------------------------------------------
	 FIELD_1="<"&"%=DETAIL_CONN(""" & TEMP_FIELD_1 & """)%"&">"
CALL TOTABLE(FIELD_1,"TD")	 
     TEMP_DETAIL_DB=FIELD_2&FIELD_1
CALL TOTABLE(TEMP_DETAIL_DB,"TR")

ADMIN_DETAIL_DB_1=ADMIN_DETAIL_DB_1&TEMP_DETAIL_DB


		
END IF

NEXT
	 
'-->> �ж��������Ƿ� ��Ψһ��	 
IF KEY_NUM <> 1 THEN              
   CALL MSG("����Ӧ��Ψһ,�Ҳ�Ӧ��Ϊ��...",2)
END IF
	 
	 
END IF




	



'=====================================================================
'--->>��ҳ�� ��ֵģ�壬�滻,����>>
'=====================================================================	

    '----->>  ��ֵģ�壬�滻,����>> ���/�༭ҳ��
	'----->>  ��ֵģ�壬�滻,����>> �б�ҳ�� 
	'----->>  ��ֵģ�壬�滻,����>> ��ϸҳ�� 
	'----->>  ��ֵģ�壬�滻,����>> ���ݿ������ļ�
	'----->>  ��ֵģ�壬�滻,����>> ��ҳ ��ҳ��
	'----->>  ��ֵģ�壬�滻,����>> ���������ļ�
	

	
	DIM PAGE_CONTENT,BACK_INFO
	
	FOR NUM=0 TO WEB_PAGE_NUM
	     PAGE_CONTENT="PAGE_"&NUM
	     PAGE_CONTENT=SESSION(PAGE_CONTENT)
		 CALL STRTOCODE(PAGE_CONTENT)
		 CALL WEB_CREATE(WEB_FILE(INT(NUM)),PAGE_CONTENT,"FILE")
		 '==>>�����ļ��ķ�����Ϣ
		 BACK_INFO=BACK_INFO&"������ => " & WEB_FILE(NUM) & "\n\n"
	NEXT	 
     '--->>������Ϣ
    BACK_INFO="ʱ  �� => " & now() & "\n\n" & BACK_INFO
    CALL MSG(BACK_INFO,2)
	
	
	
	
END IF
%>






<STYLE>
/*TABLE_DIV����ʽ(GO)*/
.TABLE_DIV{
FONT-SIZE:12PX;
WIDTH:120PX;
HEIGHT:18PX;
BACKGROUND-COLOR:#494949;
COLOR:#00CCFF;
TEXT-ALIGN:LEFT;
MARGIN-TOP:1PX;
FONT-FAMILY:"����";
}

.TABLE_DIV SPAN{
PADDING:5PX;
BACKGROUND-COLOR:#373737;
COLOR:#999999;
}

.TABLE_DIV A{
COLOR:#00CCFF;
TEXT-DECORATION:NONE;
WIDTH:120PX;
HEIGHT:18PX;
}
.TABLE_DIV A:HOVER{
COLOR:#00FF00;
TEXT-DECORATION:NONE;
WIDTH:120PX;
HEIGHT:18PX;
}
/*TABLE_DIV����ʽ(END)*/



/*FILED_DIV����ʽ(GO)*/
.FILED DIV{
FONT-SIZE:11PX;
COLOR:#CCCCCC;
PADDING:2PX;
PADDING-BOTTOM:3PX;
BORDER:0;
BORDER-BOTTOM:#000000 1PX DOTTED}

.FILED_DIV A{
FONT-SIZE:12PX;
WIDTH:102PX;
HEIGHT:18PX;
BACKGROUND-COLOR:#333333;
COLOR:#CCCCCC;
TEXT-ALIGN:LEFT;
MARGIN-TOP:1PX;
MARGIN-LEFT:16PX;
PADDING-LEFT:6PX;
FONT-FAMILY:"����";
TEXT-DECORATION:NONE;
}
.FILED_DIV A:HOVER{
COLOR:#FFFFFF;
WIDTH:102PX;
}
/*FILED_DIV����ʽ(END)*/



.TITLE_{COLOR:#CCCCCC;
FONT-SIZE:12PX;}

#DIV_BUTTOM DIV{
WIDTH:70PX;
HEIGHT:19PX;
BORDER:#000000 1PX SOLID;
BACKGROUND-COLOR:#333333;
MARGIN:1PX;
MARGIN-LEFT:0;
FLOAT:LEFT;
TEXT-ALIGN:CENTER;
}

#DIV_BUTTOM A{
PADDING-TOP:2PX;
FONT-SIZE:11PX;
WIDTH:70PX;
HEIGHT:17PX;
}


.s_input{
	width:80px;
	height:18px;
	bacKground-color:#333333;
	border:1px #555555 solid;
	color:#CCCCCC;
	margin:0;
}
.l_input{
width:305px;
height:16px;
bacKground-color:#333333;
border:1px #333333 dotted;
color:#CCCCCC;
margin:0;
}


.select1
{   width:  110px;
    height: 18px;
    overflow: hidden; /*������С���ǣ���Ϊ���Ϊ110px,��select���Ϊ130px*/
 display:block;
}
.select2{
border: 1px solid #000;
width:130px;
display:block;
height: 18px;
overflow:hidden;
background:url(images/icons/10.gif) right top no-repeat 18px 18px;   /*�������õ����µ�Сͼ��*/
}

select
{
	FONT-SIZE: 11px;
	position: relative;
	left: -2px;
	top: -2px;
	width: 130px;
	color: #000;
}
textarea{
background-color:#333333;
border:0;
color:#CCCCCC;
font-size:11px;
}


BODY {
	MARGIN-LEFT: 0PX;
	MARGIN-TOP: 0PX;
	MARGIN-RIGHT: 0PX;
	MARGIN-BOTTOM: 0PX;
	BACKGROUND-COLOR: #000000;
	SCROLLBAR-BASE-COLOR:#333333;
	SCROLLBAR-DARK-SHADOW-COLOR:#000000;
	SCROLLBAR-ARROW-COLOR:#000000;
	SCROLLBAR-SHADOW-COLOR:#333333;
	SCROLLBAR-FACE-COLOR:#333333;
	SCROLLBAR-HIGHLIGHT-COLOR:#333333;
	overflow:auto;
	color:#CCCCCC;
}


A:LINK {
	COLOR: #00CCFF;
	TEXT-DECORATION: NONE;
}
A:VISITED {
	TEXT-DECORATION: NONE;
	COLOR: #00CCFF;
}
A:HOVER {
	TEXT-DECORATION: NONE;
	COLOR: #0099FF;
}
A:ACTIVE {
	TEXT-DECORATION: NONE;
	COLOR: #0099FF;
}
.STYLE1 {COLOR: #FFFFFF}
</STYLE>

<SCRIPT language="JavaSCRIPT" type="text/javaSCRIPT"> 
function ChangeDIV(DIVId,DIVName,zDIVCount) 
{for(i=0;i<=zDIVCount;i++)
{document.getElementById(DIVName+i).style.display="none";}
 document.getElementById(DIVName+DIVId).style.display="block";} 
</SCRIPT>



<%
'=====================================================================

' --->>      ʹ��SESSION("SYS_PAGE")�ж�ҳ��'

'=====================================================================
            SYS_PAGE=SESSION("SYS_PAGE")
SELECT CASE SYS_PAGE


'=====================================================================
'-->>  ϵͳ��½ҳ��
'=====================================================================
CASE ""
%>










<TITLE><%=WEB_INFO%>�����ȵ�½...</TITLE>
<FORM ID="FORM1" NAME="FORM1" METHOD="POST" ACTION="">
<br />
<table width="100" border="0" align="center" cellspacing="15" bgcolor="#232323">
    <tr>
      <td><table width="200" border="0" align="center" cellpadding="0" cellspacing="10" bgcolor="#333333" style="font-size:11px; color:#CCCCCC; border:#666666 1px dotted;">
        <tr>
          <td height="30" colspan="2"><span class="TABLE_DIV" style="WIDTH:100%; BACKGROUND-COLOR:#313131; TEXT-ALIGN:CENTER"><strong><%=WEB_INFO%>,��½</strong></span></td>
        </tr>
        <tr>
          <td height="10" align="center" bgcolor="#232323"><input name="LOGIN_PASS" type="text" id="LOGIN_PASS" style="WIDTH:120px; HEIGHT:15PX; BACKGROUND-COLOR:#232323; BORDER:1 #666666 SOLID; COLOR:#FFFFFF" value="1"/></td>
          <td width="72" bgcolor="#232323"><input type="submit" name="BUTTON" id="BUTTON" value="��½" style="WIDTH:100%; HEIGHT:15PX;BACKGROUND-COLOR:#232323; BORDER:1 #666666 SOLID; COLOR:#FFFFFF"/></td>
        </tr>
      </table></td>
    </tr>
  </table>
</FORM>










<%
'ѡ�����ݿ�
'==================================================
CASE "SELECT_DB"
%>
<TITLE><%=WEB_INFO%>����ѡ�����ݿ�...</TITLE>
<BR />
<TABLE WIDTH="680" HEIGHT="49" BORDER="0" ALIGN="CENTER" CELLPADDING="0" CELLSPACING="1">
<FORM NAME="DB_FORM" METHOD="POST" ACTION="" STYLE="MARGIN:0">
<TR>
<TD HEIGHT="15" colspan="2" BGCOLOR="#333333">

<DIV id="DIV_BUTTOM" style="width:100%; padding-left:8PX; padding-right:8px;padding-top:10PX;">
<%=SESSION("SYS_MENU")%>


<DIV STYLE="width:100%; border:0;">
<DIV style="float:right"><a href="?DIY_PAGE=OUT">�˳�ϵͳ</a></DIV>

<DIV style="width:560px;height:15px;border:0; font-size:11px; padding-top:5px; padding-right:30px;float:right; text-align:right">
&lt;= <%=WEB_INFO%>��ϵͳ��������ģ��
</DIV>
</DIV>
</DIV>


<DIV id="BigDIV" style="border:dotted 1px #666666;width:100%;margin:8px;">












<DIV id="JKDIV_0" style="display:block;">
<%'CSS��ʽҳ%>
<textarea name="PAGE_0" rows="26" id="PAGE_0" style="width:100%">
.EDIT_INPUT_L{width:100%; height:16PX;}
.EDIT_BUTTOM {width:70PX; height:25PX; background-color:#EEEEEE}
.EDIT_TABLE  {background-color:#FFFFFF; border:1PX #999999 dotted}
.EDIT_TABLE TD{background-color:#F7F7F7;}
</textarea>
</DIV> 




<DIV id="JKDIV_1" style="display:none;">
<%'�༭�����ҳ��%>
<textarea name="PAGE_1" rows="26" id="PAGE_1" style="width:100%">
&lt;%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%&gt;
&lt;!--#INCLUDE FILE="$DATABASE_CONN$"--&gt;

&lt;%
'|---------------------------------------------------------------------------------------|
'|---------------------------   ��Ҫ����,���ձ����������ݿ����    --------------------|
'                                                                                        |
DIM EDIT_TYPE,EDIT_ID,FORM_TYPE,EORRE_INFO                              '                |
                                                                        '                |
    EDIT_TYPE  =  REQUEST.QueryString("EDIT_TYPE")                      '�ж� �༭���޸� |
    EDIT_ID    =  REQUEST.QueryString("EDIT_ID")                        '�༭�����IDֵ  |
	FORM_TYPE  =  REQUEST.Form("FORM_TYPE")                             '�ж��Ƿ��ύ��|
	EORRE_INFO =  "&lt;DIV STYLE='FONT-SIZE:12PX;'&gt;���ʳ���...&lt;/DIV&gt;"  '������Ϣ        |
	   
	'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	
DIM DB_TABL,DB_KEY                                                     '���ݿ����
    DB_TABL = "$TABLE_NAME$"
	DB_KEY  = "$TABLE_KEY$"
	
'|---------------------------------------------------------------------------------------


  SUB EORRE_BACK(STR)
      RESPONSE.Write("&lt;script&gt;alert('" & STR & "');history.go(-1);&lt;/script&gt;")
	  RESPONSE.End()
  END SUB

  SUB FORM_CHECK(S_NAME,S_STR,LIMIT)
     SELECT CASE LIMIT
  
            CASE "no"                  '�ж��Ƿ�Ϊ��
			     IF ISEMPTY(S_STR) THEN
				    CALL EORRE_BACK("������� " & S_NAME & " ,����Ϊ�գ�")
			     END IF
				 
            CASE "LIMIT_2"
			     
            CASE "LIMIT_3"
			     IF REPLACE(S_STR,"@","")=S_STR THEN
                    CALL EORRE_BACK("������� " & S_NAME & " �����ϣ���˶ԣ�")
				 END IF
                 
            CASE "LIMIT_4"
			     IF NOT ISNUMERIC(S_STR) THEN
				    CALL EORRE_BACK("������� " & S_NAME & " �������֣���˶ԣ�")
				 END IF
                
            CASE "LIMIT_5"
			     IF NOT ISDATE(S_STR) THEN
				    CALL EORRE_BACK("������� " & S_NAME & " ������Ч��ʱ�䣬��˶ԣ�")
				 END IF
     END SELECT
  END SUB

%&gt;


&lt;%
'---------------------------------------------------------------------
'�ж����޸ģ��������||| �������ݲ���                                |
'---------------------------------------------------------------------
IF NOT ISEMPTY(FORM_TYPE) THEN

'---------���ձ�����(ѭ�������)======&gt;&gt;&gt; 
          

$����������$
      
'�磺F_$�ֶ�$  =  REQUEST.Form("F_$�ֶ�$")
         
$�����ж�$


         SET EDIT_CONN=SERVER.CREATEOBJECT("ADODB.RECORDSET")
         
			 IF EDIT_TYPE="ADD" THEN
		        EDIT_CONN_STR="SELECT * FROM " & DB_TABL
			    EDIT_CONN.OPEN EDIT_CONN_STR,CONNSTR,1,3
				EDIT_CONN.ADDNEW
				BACK_INFO="=&gt; " & NOW() & ",��ӳɹ���"

		 ELSEIF EDIT_TYPE="EDIT" AND NOT ISEMPTY(EDIT_ID) AND ISNUMERIC(EDIT_ID) THEN
                EDIT_CONN_STR="SELECT * FROM " & DB_TABL & " WHERE " & DB_KEY & "=" & INT(EDIT_ID)
				EDIT_CONN.OPEN EDIT_CONN_STR,CONNSTR,1,3
				BACK_INFO="=&gt; " & NOW() & ",�޸ĳɹ���"

            ELSE
			
			    RESPONSE.Write(EORRE_INFO)
                RESPONSE.End() 	
		     END IF
 
 
'---------�жϲ�д������======&gt;&gt;&gt;
			 
			 IF NOT EDIT_CONN.EOF THEN
	
				$д��������$
				'�磺EDIT_CONN("$�ֶ�$")=F_$�ֶ�$

				RESPONSE.Write("&lt;script&gt;alert('" & BACK_INFO & "');&lt;/script&gt;")
				BACK_INFO=""
			 END IF

			EDIT_CONN.UPDATE
		    EDIT_CONN.CLOSE
		SET EDIT_CONN= NOTHING
END IF

%&gt;

&lt;%
'-----------------------------------------------------------------------
'�ж����޸ģ��������||| ��ʾ���ݶ�ȡ����
'-----------------------------------------------------------------------

    IF EDIT_TYPE="ADD" THEN
	
       'F_$�ֶ�$=REQUEST.Form("F_$�ֶ�$")     [47���Ѿ�����,���п��Ժ���]

ELSEIF EDIT_TYPE="EDIT" AND NOT ISEMPTY(EDIT_ID) AND ISNUMERIC(EDIT_ID) THEN
       SET READ_CONN=SERVER.CreateOBJect("adodb.recordset")
	       READ_CONN_STR="SELECT * FROM " & DB_TABL & " WHERE " & DB_KEY & "=" & INT(EDIT_ID)
		   READ_CONN.OPEN READ_CONN_STR,CONNSTR,1,3
		   IF NOT READ_CONN.EOF THEN

		    $��ȡ������$
		   '�磺F_$�ֶ�$=READ_CONN("$�ֶ�$")

		   ELSE
	       RESPONSE.Write(EORRE_INFO)
           RESPONSE.End() 	   
		   END IF
		   READ_CONN.CLOSE
	   SET READ_CONN=NOTHING
ELSE  
      RESPONSE.Write(EORRE_INFO)
      RESPONSE.End()   
END IF
%&gt;
&lt;link href="../STYPE/ADMIN_CSS.CSS" rel="stylesheet" type="text/css" /&gt;
&lt;form id="ADMIN_EDIT_FORM" name="ADMIN_EDIT_FORM" method="post" action="?EDIT_TYPE=&lt;%=EDIT_TYPE%&gt;&EDIT_ID=&lt;%=EDIT_ID%&gt;" onSubmit="return CheckForm();" &gt;

&lt;table width="100%" border="0" cellpadding="0" cellspacing="10" bgcolor="#F7F7F7"&gt;
&lt;tr&gt;
&lt;td&gt;

&lt;table width="100%" border="0" cellpadding="0" cellspacing="2" class="EDIT_TABLE"&gt;

$��������$

&lt;tr&gt;
&lt;td align="right"&gt;&nbsp;&lt;/td&gt;
&lt;td&gt;
&lt;input name="button" type="reset" class="EDIT_BUTTOM" id="button" value="ȡ��"&gt;
&lt;input name="button" type="submit" class="EDIT_BUTTOM" id="button" value="�ύ"&gt;
&lt;input name="FORM_TYPE" type="hidden" value="&lt;%=EDIT_TYPE%&gt;" /&gt;
&lt;/td&gt;
&lt;/tr&gt;
&lt;/table&gt;
&lt;/td&gt;
&lt;/tr&gt;
&lt;/table&gt;
&lt;/form&gt;

&lt;script&gt;
function CheckForm()
{

$���ж�$

}
&lt;/script&gt;
</textarea>
</DIV> 












<DIV id="JKDIV_2" style="display:none;">
<%'�б�ҳ��%>
<textarea name="PAGE_2" rows="26" id="PAGE_2" style="width:100%">
&lt;%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%&gt;
&lt;!--#INCLUDE FILE="$DATABASE_CONN$"--&gt;

&lt;%
'|---------------------------------------------------------------------------------------
'|---------------------------   ��Ҫ����,���ձ����������ݿ����    --------------------|
'                                                                                        |
DIM DB_TABL,DB_KEY                            '���ݿ����
DIM DEL_ID,DEL_BACK,PAGE_SIZE,LIST_NUM,LIST_CONN_NUM        'ɾ�����ص�����


'���ϵ���������Ϣ ����������������
    DB_TABL = "$TABLE_NAME$"                          '��ǰ�����ı�
	DB_KEY="$TABLE_KEY$"                           '����

	PAGE_SIZE=10                              'ÿҳ��ʾ��Ŀ
	
	DEL_ID=REQUEST.QueryString("DEL_ID")      'ɾ����IDֵ
	
SUB SHOW_MSG(STR)
    RESPONSE.Write("&lt;script&gt;alert('"&STR&"');&lt;/script&gt;")
END SUB
	
	
	
'--->>�жϲ�ɾ��
IF DEL_ID&lt;&gt;"" AND ISNUMERIC(DEL_ID) THEN
SET LIST_DEL_CONN=SERVER.CREATEOBJECT("ADODB.RECORDSET")

    LIST_DEL_CHECK_STR="SELECT * FROM " & DB_TABL & "  WHERE "&DB_KEY&"=" & INT(DEL_ID)
	LIST_DEL_CONN_STR="DELETE * FROM " & DB_TABL & "  WHERE "&DB_KEY&"=" & INT(DEL_ID)
	LIST_DEL_CONN.OPEN LIST_DEL_CHECK_STR,CONNSTR,1,3
	
	IF NOT LIST_DEL_CONN.EOF THEN
	   LIST_DEL_CONN.CLOSE
	   LIST_DEL_CONN.OPEN LIST_DEL_CONN_STR,CONNSTR,1,3
	   LIST_DEL_CONN.UPDATE
       DEL_BACK="TRUE"
	ELSE
	   LIST_DEL_CONN.CLOSE
	END IF
SET	LIST_DEL_CONN=NOTHING
END IF
	
	

'-------&gt;��ȡ����
SET LIST_CONN=SERVER.CREATEOBJECT("ADODB.RECORDSET")
    LIST_CONN_STR="SELECT * FROM " & DB_TABL & " ORDER BY "&DB_KEY&" DESC"
	LIST_CONN.OPEN LIST_CONN_STR,CONNSTR,1,3
'-------&gt;���ݷ�ҳ

IF NOT LIST_CONN.eof THEN
   LIST_CONN.PAGESIZE=PAGE_SIZE
IF REQUEST.QueryString("PAGE")&lt;&gt;"" THEN
   PAGE=cint(REQUEST.QueryString("PAGE"))   
   IF PAGE&lt;1 THEN PAGE=1
   IF PAGE&gt;LIST_CONN.PAGECOUNT THEN PAGE=LIST_CONN.PAGECOUNT
ELSE
   PAGE=1
END IF
   LIST_CONN.ABSOLUTEPAGE=PAGE
END IF

'-------&gt;��ҳ����ı���
PAGE=PAGE                      '��ǰҳ��
LIST_NUM=3                     '
'FRIST_PAGE=1                   ' ��ҳ
LAST_PAGE=LIST_CONN.PAGECOUNT  ' ���ҳ
%&gt;
&lt;table width="100%" border="0" cellpadding="0" cellspacing="1" class="LIST_TABLE"&gt;
  &lt;tr&gt;
  
  
$ADMIN_LIST_DB_1$

    
    &lt;td height="25" colspan="3" align="center"&gt;����&lt;/td&gt;
  &lt;/tr&gt;
&lt;%
FOR LIST_CONN_NUM=1 TO PAGE_SIZE
IF  LIST_CONN.EOF OR LIST_CONN.BOF THEN EXIT FOR
%&gt;
&lt;tr&gt;


$ADMIN_LIST_DB_2$


&lt;td width="60" align="center"&gt;
&lt;a href="?PAGE=&lt;%=PAGE%&gt;&DEL_ID=&lt;%=LIST_CONN(DB_KEY)%&gt;" onclick="return confirm('�Ƿ�ȷ��ɾ��?');"&gt;ɾ��&lt;/a&gt;
&lt;/td&gt;
&lt;td width="60" align="center"&gt;
&lt;a href="ADMIN_&lt;%=DB_TABL%&gt;_DETAIL.ASP?DETAIL_ID=&lt;%=LIST_CONN(DB_KEY)%&gt;"&gt;Ԥ��&lt;/a&gt;
&lt;/td&gt;
&lt;td width="60" align="center"&gt;
&lt;a href="ADMIN_&lt;%=DB_TABL%&gt;_EDIT.ASP?PAGE=&lt;%=PAGE%&gt;&EDIT_TYPE=EDIT&EDIT_ID=&lt;%=LIST_CONN(DB_KEY)%&gt;"&gt;�޸�&lt;/a&gt;
&lt;/td&gt;

&lt;/tr&gt;
&lt;%
    LIST_CONN.MOVENEXT
     NEXT
    LIST_CONN.CLOSE
SET LIST_CONN =NOTHING
%&gt;  
&lt;/table&gt;

&lt;!--#include file="ADMIN_ALL_PAGINATION.ASP"--&gt;
&lt;%
IF  DEL_BACK="TRUE" THEN
CALL SHOW_MSG("�Ѿ��ɹ�ɾ��...")
END IF
%&gt;
</textarea>
</DIV> 












<DIV id="JKDIV_3" style="display:none;">
<%'��ϸҳ��%>
<textarea name="PAGE_3" rows="26" id="PAGE_3" style="width:100%">
&lt;%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%&gt;
&lt;!--#INCLUDE FILE="$DATABASE_CONN$"--&gt;

&lt;%
'|---------------------------------------------------------------------------------------
'|---------------------------   ��Ҫ����,���ձ����������ݿ����    --------------------|
'                                                                                        |
DIM DB_TABL,DB_KEY                            '���ݿ����
DIM DEL_ID,DEL_BACK,PAGE_SIZE,DETAIL_NUM        'ɾ�����ص�����


'���ϵ���������Ϣ ����������������
    DB_TABL = "$TABLE_NAME$"           '��ǰ�����ı�
	DB_KEY="$TABLE_KEY$"                '����
    DETAIL_ID=REQUEST.QueryString("DETAIL_ID")

'��ȡ����
IF DETAIL_ID="" OR ISNUMERIC(DETAIL_ID)=FALSE THEN
   RESPONSE.Write("&lt;script&gt;alert('��������...');history.go(-1);&lt;/script&gt;")
   RESPONSE.End()
END IF

SET DETAIL_CONN=SERVER.CREATEOBJECT("ADODB.RECORDSET")
    DETAIL_CONN_STR="SELECT * FROM " & DB_TABL & "  WHERE "&DB_KEY&"=" & INT(DETAIL_ID)
	DETAIL_CONN.OPEN DETAIL_CONN_STR,CONNSTR,1,3    
	IF DETAIL_CONN.EOF THEN
	   RESPONSE.Write("&lt;script&gt;alert('��������...');history.go(-1);&lt;/script&gt;")
	   RESPONSE.End()
	END IF
%&gt;

&lt;table width="100%" border="0" cellpadding="0" cellspacing="0"&gt;

  
$ADMIN_DETAIL_DB_1$
   

&lt;/table&gt;


&lt;%
	DETAIL_CONN.CLOSE
SET DETAIL_CONN=NOTHING
%&gt;
</textarea>
</DIV> 












<DIV id="JKDIV_4" style="display:none;">
<textarea name="PAGE_4" rows="26" id="PAGE_4" style="width:100%">
����ҳ��
</textarea>
</DIV> 












<DIV id="JKDIV_5" style="display:none;">
<textarea name="PAGE_5" rows="26" id="PAGE_5" style="width:100%">
&lt;%
'===========��Ȩ��Ϣ( GO )============
'<%=WEB_INFO%>
'===========��Ȩ��Ϣ��END��============
%&gt;
&lt;%
'=============================================
'---------------(���ݿ��ע�벿��)------------
SQL_INjDATA ="'|=|EXEC|INSERT|SELECT|DELETE|UPDATE|COUNT|*|%|CHR|MID|MASTER|TRUNCATE|CHAR|DECLARE|SET"
SQL_INj = SPLIT(SQL_INjDATA,"|")
'-----------------------------------------
     IF REQUEST.QUERYSTRING<>"" THEN
        FOR EACH SQL_GET IN REQUEST.QUERYSTRING
        FOR SQL_DATA=0 TO UBOUND(SQL_INj)
            IF INSTR(REQUEST.QUERYSTRING(SQL_GET),SQL_INj(SQL_DATA))>0 THEN
                 RESPONSE.WRITE("��Ǹ�����ʳ���...")
                 RESPONSE.END
             END IF
         NEXT
         NEXT
      END IF
%&gt;
&lt;%
'=============================================
'---------------(���ݿ����Ӳ���)--------------
  DIM CONNS,CONNSTR,TIME1,TIME2,MDB
      TIME1=TIMER
      MDB="$DATA_CONN_PATH$"
   ON ERROR RESUME NEXT
      SET CONNS=SERVER.CREATEOBJECT("ADODB.CONNECTION")
          CONNSTR="DRIVER=MICROSOFT ACCESS DRIVER (*.MDB);DBQ="+SERVER.MAPPATH(MDB)
          CONNS.OPEN CONNSTR
 '------------------------------------
  IF ERR THEN
     ERR.CLEAR
     SET CONNS = NOTHING
         RESPONSE.WRITE "&lt;DIV STYLE='FONT-SIZE:12PX;'&gt;��ʱ�޷���ȡ���ݿ⣬�����~&lt;/DIV&gt;"
         RESPONSE.END
  END IF
%&gt;
</textarea>
</DIV> 


<DIV id="JKDIV_6" style="display:none;">
<textarea name="PAGE_6" rows="26" id="PAGE_6" style="width:100%">
&lt;%
'������� ==&gt;&gt;�ɵ�����б�ҳ���ṩ
'PAGE=0               '��ǰҳ��
'LIST_NUM=3           '
'FRIST_PAGE=0         ' ��ҳ
'LAST_PAGE=10         ' ���ҳ


NUM1=PAGE-LIST_NUM
NUM2=PAGE+LIST_NUM

IF NUM1&lt;1 THEN NUM1=1
IF NUM2&gt;LAST_PAGE THEN NUM2=LAST_PAGE
%&gt;

&lt;table width="100%" border="0" cellpadding="0" cellspacing="0"&gt;
  &lt;tr&gt;
    &lt;td width="69%" align="center"&gt;
&lt;div class="PAGINATION"&gt;

&lt;a href="?PAGE=1"&gt;��ҳ&lt;/a&gt;
&lt;%

FOR NUM=NUM1 TO NUM2
%&gt;
&lt;a href="?PAGE=&lt;%=NUM%&gt;"&gt;&lt;%=NUM%&gt;&lt;/a&gt;
&lt;%
NEXT
%&gt;


&lt;a href="?PAGE=&lt;%=LAST_PAGE%&gt;"&gt;βҳ[&lt;%=LAST_PAGE%&gt;]&lt;/a&gt;

&lt;/div&gt;
    &lt;/td&gt;
    &lt;td width="31%"&gt;&nbsp;&lt;/td&gt;
&lt;/tr&gt;
  
&lt;/table&gt;
</textarea>
</DIV> 









</DIV>


</TD>
</TR>

<TR>
<TD HEIGHT="25" WIDTH="100%" BGCOLOR="#666666" style="padding-left:8px;">
<input type="FILE" name="DATA_PATH" id="DATA_PATH" style="WIDTH:100%; HEIGHT:15PX; BACKGROUND-COLOR:#333333; BORDER:1PX #666666 SOLID;COLOR:#CCCCCC" />
</TD>
<TD height="25" BGCOLOR="#666666" style="padding-right:8px;">
<input type="SUBMIT" value="ȷ��" style="WIDTH:60PX; HEIGHT:15PX; BACKGROUND-COLOR:#333333; BORDER:1PX #666666 SOLID; COLOR:#CCCCCC"/>
</TD>
</TR>
</FORM>
</TABLE>










<%
'ѡ�����ݿ�
'==================================================
CASE "SHOW_DB_LIST"
%>
<TITLE><%=WEB_INFO%>�����ݿ����...</TITLE>

<%
IF ERR THEN
   ERR.CLEAR
   SET CONNS = NOTHING

'��ȡ���ݿ�ʧ�ܣ�����ѡ��ҳ��
    SESSION("SYS_PAGE")="SELECT_DB"
    RESPONSE.REDIRECT("INDEX.ASP")
    RESPONSE.END()

ELSE
%>




<TABLE WIDTH="0" HEIGHT="100%" BORDER="0" CELLPADDING="0" CELLSPACING="1">
<tr>
<td height="30" colspan="2" bgcolor="#000000" class="TITLE_" style="padding-left:10px; font-size:11px;">
    <strong>�� <%=web_info%></strong> / QQ:394716221</td>
</tr>
<TR>
  <TD VALIGN="TOP">

<%
'------------------------
SET OBJCONN=SERVER.CREATEOBJECT("ADODB.CONNECTION")
    OBJCONN.OPEN CONNSTR
'------------------------
SET RSSCHEMA=OBJCONN.OPENSCHEMA(20)
    RSSCHEMA.MOVEFIRST
'------------------------
DO UNTIL RSSCHEMA.EOF
   IF RSSCHEMA("TABLE_TYPE")="TABLE" THEN
'------------------------  
%>

<DIV CLASS="TABLE_DIV">
<A HREF='?EDIT_TABLE=<%=RSSCHEMA("TABLE_NAME")%>'>
<SPAN>+</SPAN>

  <%IF REQUEST.QUERYSTRING("EDIT_TABLE")=RSSCHEMA("TABLE_NAME") THEN%>
    <FONT STYLE="COLOR:#00FF00;"><%=RSSCHEMA("TABLE_NAME")%>&nbsp;-</FONT>
    <%ELSE%>
    <%=RSSCHEMA("TABLE_NAME")%>
  <%END IF%>
</A></DIV>


<%     
    END IF
    RSSCHEMA.MOVENEXT
    LOOP
SET OBJCONN=NOTHING
%>
</TD>
  
<TD WIDTH="100%" ROWSPAN="2" VALIGN="TOP" BGCOLOR="#333333">
<DIV STYLE="WIDTH:100%; HEIGHT:19PX; BACKGROUND-COLOR:#242424;" ID="DIV_BUTTOM">
<DIV><A HREF="?DIY_PAGE=RESELECT">��ѡ</A></DIV>
<DIV><A HREF="?DIY_PAGE=OUT">�˳�</A></DIV>

</DIV>
<DIV CLASS="FILED" ID="FILED">
<%
   EDIT_TABLE=REQUEST.QueryString("EDIT_TABLE")
IF EDIT_TABLE<>"" THEN
%>

<DIV >
<form id="form1" NAME="form1" method="post" action="?EDIT_TABLE=<%=TABLE_NAME%>">
  <table width="850" border="0" align="center" cellpadding="0" cellspacing="8" bgcolor="#272727">
    <tr>
      <td align="center">

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#181818" style="font-size:11px; color:#CCCCCC; border:#666666 1px dotted;">
        <tr>
          <td width="40" height="20" align="center" bgcolor="#202020">ѡ��          </td>
          <td width="4%" align="center" bgcolor="#202020">���</td>
          <td align="center" bgcolor="#202020">�ֶ�</td>
          <td align="center" bgcolor="#202020">�ֶ�����</td>
          <td width="60" align="center" bgcolor="#202020">ʹ������</td>
          <td align="center" bgcolor="#202020">��󳤶�</td>
          <td align="center" bgcolor="#202020">�����/����</td>
          <td align="center" bgcolor="#202020">�����ַ��޶�</td>
          <td align="center" bgcolor="#202020">��ֵ/Ψһֵ</td>
          <td width="50" align="center" bgcolor="#202020">JS���</td>
        </tr>
        
<%
SET SHOW_FILED_CONN = SERVER.CREATEOBJECT("ADODB.RECORDSET")
    SHOW_FILED_CONN_STR = "SELECT * FROM " & EDIT_TABLE
    SHOW_FILED_CONN.OPEN SHOW_FILED_CONN_STR,CONNSTR,0,1
	
	FOR I=0 TO SHOW_FILED_CONN.FIELDS.COUNT-1
%>

        <tr>
          <td align="center" bgcolor="#333333">
<input name="FIELD_<%=I%>" type="checkbox" id="FIELD_<%=I%>" value="<%=SHOW_FILED_CONN.FIELDS.ITEM(I).NAME%>" checked="checked" />
</td>
          <td align="center" bgcolor="#333333">
          <%=I%>
          </td>
          <td bgcolor="#333333" style="padding-left:5px;">
		  <%=SHOW_FILED_CONN.FIELDS.ITEM(I).NAME%>
          </td>
<td bgcolor="#333333">
  <input name="FIELD_SHOW_<%=I%>" type="text" class="s_input" id="FIELD_SHOW_<%=I%>" style="width:100%" value='<%=SHOW_FILED_CONN.FIELDS.ITEM(I).NAME%>' />  </td>
          <td align="center" bgcolor="#333333">
          <input name="FIELD_TYPE_<%=I%>" type="hidden" id="FIELD_TYPE_<%=I%>" value="<%=SHOW_FILED_CONN.FIELDS.ITEM(I).TYPE%>" />
          <%CALL GET_TYPE(CSTR(SHOW_FILED_CONN.FIELDS.ITEM(I).TYPE))%>          </td>
          <td align="center" bgcolor="#333333">
          <input name="FIELD_MAX_<%=I%>" type="text" class="s_input" id="FIELD_MAX_<%=I%>" style="width:100%" value="255"/></td>

<td width="80px" align="center" bgcolor="#333333">
<DIV style="height:13px; overflow: hidden; border: 0;width:75px;"> 
<DIV style="position:absolute; left:-3px; top:-3px;clip: rect(2 108 19 2);width:75px;"> 
<select name="FIELD_INPUT_TYPE_<%=I%>" class="s_input">
<option value="text">�ı���</option>
<option value="textarea">���п�</option>
<option value="file">�ı���</option>
<option value="password">�����</option>
<option value="radio">��ѡ��</option>
<option value="checkbox">��ѡ��</option>
<option value="hidden">���ؿ�</option>
<option value="html">��̬��</option>
<option value="real_ip">�� IP</option>
<option value="add_1">�ۼ�+1</option>
<option value="html">��̬��</option>
</select>
</DIV>
</DIV>
</td>
<td width="80" align="center" bgcolor="#333333">
<DIV style="height:13px; overflow: hidden; border: 0;width:75px;">  
<DIV style="position:absolute; left:-3; top:-3;clip: rect(2 108 19 2);width:75px;"> 
<select name="FIELD_INPUT_LIMIT_<%=I%>" class="s_input" id="FIELD_INPUT_LIMIT_<%=I%>">
<option value="LIMIT_1">������ֶ�</option>
<option value="LIMIT_2">��26����ĸ</option>
<option value="LIMIT_3">���ʼ���ַ</option>
<option value="LIMIT_4">����������</option>
<option value="LIMIT_5">������ʱ��</option>
</select>
</DIV>
</DIV>          </td>
          <td width="80" align="center" bgcolor="#333333">
          
<DIV style="height:13px; overflow: hidden; border: 0;width:75px;"> 
<DIV style="position:absolute; left:-3; top:-3;clip: rect(2 108 19 2);width:75px;">           
<select name="FIELD_INPUT_EMPEY_TYPE_<%=I%>" class="s_input" id="FIELD_INPUT_EMPEY_TYPE_<%=I%>">
<option value='<%=SHOW_FILED_CONN.FIELDS.ITEM(I).NAME%>'>����ֵ</option>
<option value="yes">�ɿ�ֵ</option>
<option value="no" selected="selected">����ֵ</option>
</select>
</DIV></DIV> </td>
          <td align="center" bgcolor="#333333">
          <input name="FIELD_INPUT_CHECK_TYPE_<%=I%>" type="checkbox" id="FIELD_INPUT_CHECK_TYPE_<%=I%>" checked="checked" /></td>
        </tr>

<%
	NEXT
	
	SHOW_FILED_CONN.CLOSE
SET	SHOW_FILED_CONN=NOTHING
%>   
        
        <tr>
          <td align="center" bgcolor="#202020">
          <input type="hidden" name="DB_MAX" value="<%=I%>" />
          <input type="hidden" name="DB_FORM" value="YES" />
          </td>
          <td align="center" bgcolor="#202020">&nbsp;</td>
          <td rowspan="2" align="center" bgcolor="#202020">
          <input name="button3" type="reset" class="s_input" id="button3" value="ȡ��" /></td>
          <td rowspan="2" align="center" bgcolor="#202020">
          <input name="button2" type="submit" class="s_input" id="button2" value="�ύ" />
          </td>
          <td align="center" bgcolor="#202020">&nbsp;</td>
          <td align="center" bgcolor="#202020">&nbsp;</td>
          <td align="center" bgcolor="#202020">&nbsp;</td>
          <td align="center" bgcolor="#202020">&nbsp;</td>
          <td align="center" bgcolor="#202020">&nbsp;</td>
          <td align="center" bgcolor="#202020">&nbsp;</td>
        </tr>
        <tr>
          <td align="center" bgcolor="#202020">&nbsp;</td>
          <td align="center" bgcolor="#202020">&nbsp;</td>
          <td align="center" bgcolor="#202020">&nbsp;</td>
          <td align="center" bgcolor="#202020">&nbsp;</td>
          <td align="center" bgcolor="#202020">&nbsp;</td>
          <td align="center" bgcolor="#202020">&nbsp;</td>
          <td align="center" bgcolor="#202020">&nbsp;</td>
          <td align="center" bgcolor="#202020">&nbsp;</td>
        </tr>
      </table>
      </td>
    </tr>
  </table>
</form>


</DIV>

<DIV style="padding:18px; line-height:18px;">
=>
[������]��<SPAN CLASS="STYLE1">CONN_<%=WEB_DATE%>.ASP</SPAN> [-] 
���ô��룺<SPAN CLASS="STYLE1">&LT;!--#INCLUDE FILE=&QUOT;CONN_<%=WEB_DATE%>.ASP&QUOT;--&GT;</SPAN>
 <br />
 


</DIV>

<%END IF%>
</DIV></TD>
</TR>
<TR>
  <TD HEIGHT="40" VALIGN="TOP" BGCOLOR="#181818">&nbsp;</TD>
  </TR>

<%END IF%>
</TABLE>


<%
'==============================
'���� ҳ��ɸѡ
'==============================
END SELECT
%>
