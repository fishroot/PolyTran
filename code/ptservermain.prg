#DEFINE DDE_SERVICE 'POLYTRAN'
#DEFINE DEBUG_MODE  .F.

SET DEFAULT   TO JUSTPATH(SYS(16))
SET DATE      TO GERMAN
SET PROCEDURE TO 'prg\ptServerFuncs'
SET PROCEDURE TO 'prg\ptServerSession' ADDITIVE
SET CLASSLIB  TO 'classes\ptServer'

* Start Server
IF ADDPROPERTY(_Screen, 'Server', CREATEOBJECT('ptServer', DDE_SERVICE))
  WITH _Screen.Server
    .mp_DDEHandler = 'mm_ptDDEHandler'
    .mp_ExeFile    = SYS(16)
  ENDWITH
ENDIF

* Start SysTray
IF ADDPROPERTY(_Screen, 'SysTray', CREATEOBJECT('ptSysTray'))
  WITH _Screen.SysTray
    .IconFile = 'Bitmaps\Translator.ico'
    .TipText  = 'PolyTran'
    .MenuText = 'mm_ptSysTrayMenu'
    .AddIconToSystray()
  ENDWITH
ENDIF

SYS(2340, 1)
READ EVENTS

* mm_ptDDEHandler

FUNCTION mm_ptDDEHandler
LPARAMETERS p_Channel, p_Type, p_Item, p_Data, p_Format, p_Advise

  LOCAL m_Data, m_Session, m_RetVal
  
  DDEENABLED(p_Channel, .F.)
  
  WITH _SCREEN.Server
    m_Data    = IIF(VARTYPE(p_Data) = 'C', CHRTRAN(p_Data, CHR(0), ''), '')
    m_Session = .mm_GetSessionId(p_Channel)
    m_RetVal  = .F.
    
    .mm_DDELog(p_Channel, p_Type, p_Item, p_Data)
    
    DO CASE
      CASE p_Type = 'INITIATE'
        m_RetVal = .mm_CreateSession(m_Session, m_Data)
      CASE VARTYPE(.Session(m_Session)) # 'O'
      CASE p_Type = 'POKE'
        m_RetVal = .Session(m_Session).mm_Poke(p_Item, m_Data)
      CASE p_Type = 'REQUEST'
        m_RetVal = DDEPOKE(p_Channel, p_Item, .Session(m_Session).mm_Request(p_Item))
      CASE p_Type = 'EXECUTE'
        m_RetVal = .Session(m_Session).mm_Execute(m_Data)
      CASE p_Type = 'TERMINATE'
        m_RetVal = .mm_RemoveSession(m_Session)
    ENDCASE
  ENDWITH
  
  DDEENABLED(p_Channel, .T.)
  
  RETURN m_RetVal
ENDFUNC

* mm_ptSysTrayMenu

FUNCTION mm_ptSysTrayMenu
LPARAMETERS p_1, p_2, p_3, p_4

  * If the menu does have submenus, use PARAMETERS instead of LPARAMETERS
  * to allow the variables to be seen by the submenus.

  DEFINE POPUP SysTray SHORTCUT RELATIVE FROM MROW(),MCOL()
  DEFINE BAR 1 OF SysTray PROMPT 'Einstellungen'
  DEFINE BAR 2 OF SysTray PROMPT '\-'
  DEFINE BAR 3 OF SysTray PROMPT 'Beenden'

  ON SELECTION BAR 1 OF SysTray EXECSCRIPT( ;
    '_SCREEN.Server.mm_RunForm("forms\ptServerConfig")')
  ON SELECTION BAR 3 OF SysTray EXECSCRIPT( ;
    '_SCREEN.Server = .F.' + CHR(13) + ;
    '_SCREEN.SysTray = .F.' + CHR(13) + ;
    'CLEAR EVENTS' + CHR(13) + ;
    'CANCEL')

  ACTIVATE POPUP SysTray
ENDFUNC