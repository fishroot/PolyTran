* Class ptServer

DEFINE CLASS ptServer AS CUSTOM
  PROTECTED Channel(1)
  DIMENSION Session(1)
  
  mp_DDEService = ''
  mp_DDEHandler = ''
  mp_ExeFile    = ''
  mp_LogFile    = ''
  mp_LogEnabled = .F.
  mp_StartTime  = NULL
  
  * Init( cDDEService )
  
  FUNCTION Init
  LPARAMETERS p_DDEService
    
    LOCAL mo_Config, m_DDEService
    
    m_DDEService = IIF(VARTYPE(p_DDEService) = 'C', p_DDEService, JUSTSTEM(SYS(16)))
  
    DDESETSERVICE(m_DDEService, 'DEFINE')
    
    DDESETSERVICE(m_DDEService, 'POKE',    .T.)
    DDESETSERVICE(m_DDEService, 'REQUEST', .T.)
    DDESETSERVICE(m_DDEService, 'EXECUTE', .T.)
    DDESETSERVICE(m_DDEService, 'ADVISE',  .F.)
    
    IF FILE('Config\AllUsers.ini')
      mo_Config = mm_IniGet('Config\AllUsers.ini')   
    ELSE
      mo_Config = mm_IniGet('Config\Default.ini', 'allusers')
      mm_IniSave('config\allusers.ini', mo_Config)
    ENDIF
    
    WITH This
      .mp_LogFile    = mo_Config.LogFile
      .mp_LogEnabled = mo_Config.LogEnabled
      .mp_DDEService = m_DDEService
      .mp_StartTime  = DATETIME()
    ENDWITH
    
    BINDEVENT(This, 'mp_DDEHandler', This, 'mm_SetDDEHandler', 1)
  ENDFUNC
  
  * mm_CreateSession( nSession[, eData] )
  
  FUNCTION mm_CreateSession
  LPARAMETERS p_Session, p_Data
    
    IF VARTYPE(p_Session) # 'N'
      RETURN .F.
    ENDIF
    
    This.Session(p_Session) = CREATEOBJECT('ptSession', p_Data)
    
    RETURN .T.
  ENDFUNC
  
  * mm_RemoveSession( nSession )
  
  FUNCTION mm_RemoveSession
  LPARAMETERS p_Session
  
    LOCAL m_RetVal
  
    IF VARTYPE(p_Session) # 'N'
      RETURN .F.
    ENDIF
    
    WITH This
      .Session(p_Session) = .F.
      .Channel(p_Session) = .F.
      m_RetVal = VARTYPE(.Session(p_Session)) = 'L'
    ENDWITH
    
    RETURN m_RetVal
  ENDFUNC
  
  * mm_GetSessionID( nChannel )
  
  FUNCTION mm_GetSessionID
  LPARAMETERS p_Channel
  
    LOCAL m_Session
    
    IF VARTYPE(p_Channel) # 'N'
      RETURN -1
    ENDIF
    
    WITH This
      m_Session = ASCAN(.Channel, p_Channel)
    
      IF m_Session = 0
        m_Session = ASCAN(.Channel, .F.)
        
        IF m_Session = 0
          m_Session = ALEN(.Channel) + 1
          DIMENSION .Channel(m_Session), .Session(m_Session)
        ENDIF
        
        .Channel(m_Session) = p_Channel
      ENDIF
    ENDWITH
    
    RETURN m_Session
  ENDFUNC
  
  * mm_GetChannelID( nSession )
  
  FUNCTION mm_GetChannelID
  LPARAMETERS p_Session
  
    LOCAL m_Channel
    
    IF VARTYPE(p_Session) # 'N'
      RETURN -1
    ENDIF
    
    TRY
      m_Channel = This.Channel(p_Session)
    CATCH
      m_Channel = -1
    ENDTRY
    
    RETURN m_Channel
  ENDFUNC
  
  * mm_SetDDEHandler()
  
  FUNCTION mm_SetDDEHandler
    IF VARTYPE(This.mp_DDEHandler) = 'C'
      DDESETTOPIC(This.mp_DDEService, '', This.mp_DDEHandler)
    ENDIF
  ENDFUNC
  
  * mm_DDELog( nChannel, cType, cItem, cData )
  
  FUNCTION mm_DDELog
  LPARAMETERS p_Channel, p_Type, p_Item, p_Data
  
    LOCAL m_LogText
    
    IF NOT (This.mp_LogEnabled OR FILE(This.mp_LogFile))
      RETURN .F.
    ENDIF
    
    m_LogText = '[' + TRANSFORM(p_Channel) + '] ' + ;
      TRANSFORM(p_Type) + '(' + p_Item + ;
      IIF(SIGN(LEN(p_Item)) + SIGN(LEN(p_Data)) = 2, ', ', '') + ;
      TRANSFORM(p_Data) + ')' + CHR(13) + CHR(10)
    
    RETURN 0 # STRTOFILE(m_LogText, This.mp_LogFile, 1)
  ENDFUNC
  
  * mm_RunForm( cForm )
  
  FUNCTION mm_RunForm
  LPARAMETERS p_Form
  
    IF VARTYPE(p_Form) # 'C'
      RETURN .F.
    ENDIF
    
    DO FORM (p_Form)
    RETURN .T.
  ENDFUNC
  
  * Destroy()
  
  FUNCTION destroy
    DDESETSERVICE(This.mp_DDEService, 'RELEASE')
    UNBINDEVENTS(This)
  ENDFUNC
ENDDEFINE

* Class ptSession
*
*   Layer between DDE Interface and Translator Objects, which handles:
*   DataSession, Users & Configuration

DEFINE CLASS ptSession AS SESSION
  mp_IniFile    = ''
  mp_User       = ''
  mo_Translator = NULL
  
  * Init( [cItem] )

  FUNCTION init
  LPARAMETERS p_User
  
    This.mo_Translator = CREATEOBJECT('ptTranslator', ;
      mm_IniGet('Config\AllUsers.ini', NULL, 'ip_'))
    
    IF VARTYPE(p_User) = 'C'
      This.im_SetUser(p_User)
    ENDIF
    
  ENDFUNC

  * mm_Request( cItem )
  
  FUNCTION mm_Request
  LPARAMETERS p_Item

    LOCAL m_Item, m_RetVal

    IF VARTYPE(p_Item) # 'C' OR EMPTY(p_Item)
      RETURN ''
    ENDIF
   
    m_RetVal = '' 
    m_Item   = LOWER(p_Item)
    
    DO CASE
      CASE INLIST(m_Item + ' ', 'translation ', 'services ', 'source ', 'target ', 'log ')
        m_RetVal = EVALUATE('this.mo_Translator.mm_get' + m_Item + '()')
      CASE m_Item == 'progress'
        m_RetVal = GETWORDNUM(This.mo_Translator.mm_getStatus(), 1, '|')
      CASE m_Item == 'user'
        m_RetVal = This.mp_User
      CASE m_Item == 'config'
        m_RetVal = mm_ObjectToXML(This.mo_Translator, 'mp_', 'UTF8')
      CASE m_Item == 'profiles'
        m_RetVal = This.im_GetProfiles()
    ENDCASE
    
    RETURN m_RetVal
  ENDFUNC

  * mm_Poke( cItem, cData )
  
  FUNCTION mm_Poke
  LPARAMETERS p_Item, p_Data

    LOCAL m_Item, m_RetVal
    
    IF VARTYPE(p_Item) # 'C' OR EMPTY(p_Item)
      RETURN .F.
    ENDIF
    
    m_RetVal = .F.
    m_Item   = LOWER(p_Item)
    
    DO CASE
      CASE m_Item == 'user'
        m_RetVal = This.im_SetUser(p_Data)
      CASE m_Item == 'config'
        m_RetVal = This.mo_Translator.mm_SetOption(mm_XMLToObject(p_Data, '', 'UTF8'))
      OTHERWISE
        m_RetVal = This.mo_Translator.mm_SetOption(m_Item, p_Data)      
    ENDCASE
    
    RETURN m_RetVal
  ENDFUNC

  * mm_Execute( cData )
  
  FUNCTION mm_Execute
  LPARAMETERS p_Data

    IF VARTYPE(p_Data) # 'C'
      RETURN .F.
    ENDIF
    
    LOCAL m_Cmd, m_RetVal
    
    m_RetVal = .F.
    m_Cmd    = LOWER(GETWORDNUM(p_Data, 1))
       
    DO CASE
      CASE m_Cmd == 'save'
        m_RetVal = NOT EMPTY(mm_IniSave(This.mp_IniFile, ;
          This.mo_Translator, 'Translator', 'mp_'))
      CASE m_Cmd == 'reset'
        m_RetVal = This.im_SetUser(This.mp_User)
      CASE m_Cmd == 'default'
        DELETE FILE ('config\' + ALLTRIM(This.mp_User) + '.ini')
        m_RetVal = This.im_SetUser(This.mp_User)
      CASE m_Cmd == 'clean'
        m_RetVal = This.mo_Translator.mm_CleanUp()
    ENDCASE
    
    RETURN m_RetVal
  ENDFUNC

  * im_SetUser( cUser )
  
  PROTECTED FUNCTION im_SetUser
  LPARAMETERS p_User

    LOCAL m_IniFile, mo_Config
    
    m_IniFile = 'config\' + ALLTRIM(p_User) + '.ini'
    
    IF NOT FILE(m_IniFile)
      mo_Config = mm_IniGet('config\default.ini', 'default', 'mp_')
      mm_IniSave(m_IniFile, mo_Config, 'translator', 'mp_')
    ELSE
      mo_Config = mm_IniGet(m_IniFile, 'translator', 'mp_')
    ENDIF
    
    WITH This
      .mp_User    = ALLTRIM(p_User)
      .mp_IniFile = m_IniFile
      .mo_Translator.mm_SetOption(mo_Config)
    ENDWITH
    
    RETURN .T.
  ENDFUNC

  * im_GetProfiles()
  
  PROTECTED FUNCTION im_GetProfiles

    LOCAL a_Files(1), i, m_File, m_CurSelect, m_CrsProfiles, m_RetVal
    
    m_CurSelect = SELECT()
    m_CrsProfiles = SYS(2015)
    
    CREATE CURSOR (m_CrsProfiles) (Profile M)
    INDEX ON LEFT(Profile, 5) TAG SortOrder
    
    FOR i = 1 TO ADIR(a_Files, ADDBS(JUSTPATH(This.mp_IniFile)) + '*.ini')
      m_File = LOWER(FILETOSTR(ADDBS(JUSTPATH(This.mp_IniFile)) + a_Files(i, 1)))
      IF NOT '[translator]' $ m_File
        LOOP
      ENDIF
      
      INSERT INTO (m_CrsProfiles) VALUES (FORCEEXT(a_Files(i, 1), ''))
    ENDFOR
    
    CURSORTOXML(ALIAS(), 'm_RetVal', 3, 49)
    
    USE IN SELECT()
    SELECT (m_CurSelect)
    
    RETURN m_RetVal
  ENDFUNC
ENDDEFINE