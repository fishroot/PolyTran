* Add Translator to _Screen
_Screen.AddProperty('Translator', NEWOBJECT('ptClientFrmSet', 'classes\ptClient'))

IF VARTYPE(_Screen.Translator) # 'O'
  REMOVEPROPERTY(_Screen, 'Translator')
  CANCEL
ENDIF

* Add Forms to Translator Object
WITH _SCREEN.Translator
  .mm_AddForm('forms\ptClientConfig', 'frmConfigure')
  .mm_AddForm('forms\ptClientResult', 'frmResult')
ENDWITH

* Add Menu to SysMenu
SET SYSMENU AUTOMATIC
DEFINE PAD _Translator OF _MSYSMENU PROMPT 'Translator' COLOR SCHEME 3 KEY ALT+T, '^T'
ON PAD _Translator OF _MSYSMENU ACTIVATE POPUP ptMenu

DEFINE POPUP ptMenu MARGIN RELATIVE SHADOW COLOR SCHEME 4
DEFINE BAR 1 OF ptMenu PROMPT '\<Einstellungen'
DEFINE BAR 2 OF ptMenu PROMPT '\-'
DEFINE BAR 3 OF ptMenu PROMPT '\<Toolbar' KEY CTRL+T, 'Ctrl+T'
DEFINE BAR 4 OF ptMenu PROMPT '\-'
DEFINE BAR 5 OF ptMenu PROMPT 'E\<xit'

SET MARK OF BAR 3 OF ptMenu TO _Screen.Translator.frmToolbar.Visible

ON SELECTION BAR 1 OF ptMenu _Screen.Translator.frmConfigure.Visible = ;
  !_Screen.Translator.frmConfigure.Visible
ON SELECTION BAR 3 OF ptMenu EXECSCRIPT( ;
  '_Screen.Translator.frmToolbar.Visible = !_Screen.Translator.frmToolbar.Visible' + CHR(13) + ;
  'SET MARK OF BAR 3 OF ptMenu TO _Screen.Translator.frmToolbar.Visible')
ON SELECTION BAR 5 OF ptMenu EXECSCRIPT( ;
  '_Screen.Translator = .F.' + CHR(13) + ;
  'REMOVEPROPERTY(_Screen, "Translator")' + CHR(13) + ;
  'ON KEY LABEL CTRL+T' + CHR(13) + ;
  'RELEASE PAD _Translator OF _MSYSMENU')