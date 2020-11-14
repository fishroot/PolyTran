* mm_ObjectToXML( cObject[, cPrefix[, cFormat]] )
*
*   oObject: Object which contains the Attributes
*   cPrefix: Return only Attributes with that Prefix
*   cFormat: Specifies the XML format (UTF8, DBCS), Default: DBCS

FUNCTION mm_ObjectToXML
LPARAMETERS p_Object, p_Prefix, p_Format

  LOCAL m_Prefix, m_CrsXML, m_curSelect, a_Keys(1), m_Key, m_RetVal
  
  IF VARTYPE(p_Object) # 'O'
    RETURN ''
  ENDIF
  
  m_Prefix = IIF(VARTYPE(p_Prefix) # 'C', '', p_Prefix)
  m_Format = IIF(VARTYPE(p_Format) # 'C', '', p_Format)
  
  m_CrsXML = SYS(2015)
  m_CurSelect = SELECT()
  
  CREATE CURSOR (m_CrsXML) (Name C(50), Type C(1), Value M)
  AMEMBERS(a_Keys, p_Object, 0, 'U')
  FOR EACH m_Key IN a_Keys
    IF LOWER(LEFT(m_Key, LEN(m_Prefix))) # m_Prefix
      LOOP
    ENDIF
    
    INSERT INTO (m_CrsXML) (Name, Type, Value) VALUES (SUBSTR(m_Key, LEN(m_Prefix) + 1),;
      TYPE('p_Object.' + m_key), TRANSFORM(EVALUATE('p_Object.' + m_key)))
  ENDFOR
  CURSORTOXML(m_CrsXML, 'm_RetVal', 3, IIF(m_Format # 'UTF8', 0, 48))
  USE IN (m_CrsXML)
  SELECT (m_CurSelect)
  
  RETURN m_RetVal
ENDFUNC

* mm_XMLToObject( cXML[, cPrefix[, cFormat]] )
*
*   cXML: XML formated String to
*   cPrefix: Prefix added to the Attributes in the returned Object
*   cFormat: Specifies the XML format (UTF8, DBCS), Default: DBCS

FUNCTION mm_XMLToObject
LPARAMETERS p_XML, p_Prefix, p_Format

  LOCAL m_XML, m_Prefix, m_Format, m_RetVal, m_CurSelect, m_CrsXML

  m_XML    = IIF(VARTYPE(p_XML) # 'C', '', p_XML)
  m_Prefix = IIF(VARTYPE(p_Prefix) # 'C', '', p_Prefix)
  m_Format = IIF(VARTYPE(p_Format) # 'C', 'DBCS', p_Format)
  
  m_RetVal = CREATEOBJECT('empty')
  m_CrsXML = SYS(2015)
  m_CurSelect = SELECT()
  
  IF m_Format = 'UTF8'
    m_XML = STRCONV(m_XML, 11)
  ENDIF
  
  IF NOT ('<' $ m_XML) OR EMPTY(XMLTOCURSOR(m_XML, m_CrsXML))
    RETURN m_RetVal
  ENDIF

  SELECT (m_CrsXML)

  SCAN
    m_Value = IIF(Type = 'C', TRIM(Value), EVALUATE(Value))
    TRY
      ADDPROPERTY(m_RetVal, m_Prefix + TRIM(Name), m_Value)
    CATCH
    ENDTRY
  ENDSCAN
  
  USE IN SELECT()
  SELECT (m_CurSelect)
  
  RETURN m_RetVal 
ENDFUNC

* mm_IniGet( cFile[, cHeader[, cPrefix]] )
*
*   cFile: Filename of the IniFile
*   cHeader: Section in the IniFile (default: all)
*   cPrefix: Prefix added to the Attributes in the returned Object

FUNCTION mm_IniGet
LPARAMETERS p_File, p_Header, p_Prefix, p_Flags
  
  LOCAL m_Header, m_Prefix, m_Section, m_Line, m_Key, m_Value
  LOCAL a_Lines(1), mo_Object, m_This

  mo_Object = CREATEOBJECT('empty')

  IF VARTYPE(p_File) # 'C' OR NOT FILE(p_File)
    RETURN mo_Object
  ENDIF
  
  m_Header = IIF(VARTYPE(p_Header) # 'C', '', p_Header)
  m_Prefix = IIF(VARTYPE(p_Prefix) # 'C', '', p_Prefix)
  
  ALINES(a_Lines, FILETOSTR(p_File))
  
  m_Section = ''
 
  FOR EACH m_Line IN a_Lines
    m_Line = ALLTRIM(m_Line)
    
    IF m_Line = '#' OR NOT ('=' $ m_Line OR m_Line = '[')
      LOOP
    ENDIF
    
    IF STRTRAN(m_Line, STREXTRACT(m_Line, '[', ']'), '', 1, 1, 1) = '[]'
      m_Section = CHRTRAN(STREXTRACT(m_Line, '[', ']'), ' ', '_')
      LOOP
    ENDIF
    
    m_This = 'mo_Object' + IIF(EMPTY(m_Section) OR NOT EMPTY(m_Header), '', '.' + m_Section)
    
    IF EMPTY(m_Header)
      IF TYPE(m_this) = 'U'
        ADDPROPERTY(mo_Object, m_Section, CREATEOBJECT('empty'))
      ENDIF
    ELSE
      IF LOWER(m_Section) # LOWER(m_Header)
        LOOP
      ENDIF
    ENDIF
    
    m_Key   = CHRTRAN(TRIM(STREXTRACT(m_Line, '', '=')), ' ', '_')
    m_Value = LTRIM(STREXTRACT(m_Line, '='))
    
    IF EMPTY(m_Key) OR EMPTY(m_Value)
      LOOP
    ENDIF
    
    m_Value = IIF(LOWER(m_Value) == 'yes', '.T.', IIF(LOWER(m_Value) == 'no', '.F.', m_Value))
    
    TRY
      ADDPROPERTY(EVALUATE(m_This), m_Prefix + m_Key, ;
        IIF(TYPE(m_Value) $ 'UO', m_Value, EVALUATE(m_Value)))
    CATCH
    ENDTRY
  ENDFOR
  
  RETURN mo_Object
ENDFUNC

* mm_IniSave( cFile, oObject[, cHeader[, cPrefix[, nFlags]]] )
*
*   cFile: Filename of the IniFile
*   oObject: Object which contains the Attributes (and Level 1 SubObjects)
*   cHeader: Section in the IniFile (default: all)
*   cPrefix: Prefix to cut from the Attributes in the Object
*   nFlags: Method to handle existing Files (0 = Merge, 1 = OverWrite)

FUNCTION mm_IniSave
LPARAMETERS p_File, po_Object, p_Header, p_Prefix, p_Flags

  #DEFINE QUIET .T.

  LOCAL mo_Object, m_Header, m_Prefix, a_Sections(1), m_Section, m_ObjectSection, m_FileSection
  LOCAL a_Keys(1), m_Key, m_ObjectKey, m_FileKey, m_FileString, m_Value, m_Flags, m_RetVal

  IF VARTYPE(p_File) # 'C' OR VARTYPE(po_Object) # 'O'
    RETURN .F.
  ENDIF
  
  m_Header = IIF(VARTYPE(p_Header) # 'C', '', ALLTRIM(p_Header))
  m_Prefix = IIF(VARTYPE(p_Prefix) # 'C', '', ALLTRIM(p_Prefix))
  m_Flags  = IIF(VARTYPE(p_Flags) # 'N', 0, p_Flags % 2)

  IF BITTEST(m_Flags, 0)
    IF EMPTY(m_Header)
      mo_Object = po_Object
    ELSE
      mo_Object = CREATEOBJECT('empty')
      ADDPROPERTY(mo_Object, m_Header, po_Object)
    ENDIF
  ELSE
    * Get IniFile
    mo_Object = mm_IniGet(p_File, '', m_Prefix)
    
    * Merge Properties with IniFile
    IF AMEMBERS(a_Sections, po_Object) = 0
      RETURN .F.
    ENDIF
    
    DIMENSION a_Sections(ALEN(a_Sections) + 1)
    AINS(a_Sections, 1)
    
    FOR EACH m_Section IN a_Sections
      m_ObjectSection = 'po_Object' + IIF(EMPTY(m_Section), '', '.' + m_Section)
      m_FileSection   = 'mo_Object' + ;
        IIF(EMPTY(m_Header), IIF(EMPTY(m_Section), '', '.' + m_Section), '.' + m_Header)
      
      IF TYPE(m_ObjectSection) # 'O'
        LOOP
      ENDIF
      
      IF TYPE(m_FileSection) = 'U'
        ADDPROPERTY(mo_Object, IIF(EMPTY(m_Header), m_Section, m_Header), CREATEOBJECT('empty'))
      ENDIF
      
      IF AMEMBERS(a_Keys, EVALUATE(m_ObjectSection), 0, 'U') = 0
        LOOP
      ENDIF
      
      FOR EACH m_Key IN a_Keys
        IF LEFT(m_Key, LEN(m_Prefix)) # UPPER(m_Prefix)
          LOOP
        ENDIF
      
        m_ObjectKey = m_ObjectSection + '.' + m_Key
        m_FileKey   = m_FileSection + '.' + m_Key
      
        IF TYPE(m_ObjectKey) $ 'OU' OR;
          TYPE(m_ObjectKey) = TYPE(m_FileKey) AND;
          EVALUATE(m_ObjectKey) == EVALUATE(m_FileKey)
          LOOP
        ENDIF
        
        TRY
          REMOVEPROPERTY(EVALUATE(m_FileSection), m_Key)
          ADDPROPERTY(EVALUATE(m_FileSection), m_Key, EVALUATE(m_ObjectKey))
        CATCH
        ENDTRY
      ENDFOR
    ENDFOR
  ENDIF

  * Create IniFile
  AMEMBERS(a_Sections, mo_Object)
  DIMENSION a_Sections(ALEN(a_Sections) + 1)
  AINS(a_Sections, 1)

  m_FileString = ''
  FOR EACH m_Section IN a_Sections
    m_FileSection = 'mo_Object' + IIF(EMPTY(m_Section), '', '.' + m_Section)
    
    IF TYPE(m_FileSection) # 'O'
      LOOP
    ENDIF
    
    m_FileString = m_FileString +;
      IIF(EMPTY(m_FileString), '', CHR(13) + CHR(10)) +;
      IIF(EMPTY(m_Section), '', '[' + m_Section + ']' + CHR(13) + CHR(10))

    AMEMBERS(a_Keys, EVALUATE(m_FileSection))
    
    FOR EACH m_Key IN a_Keys
      IF TYPE(m_FileSection + '.' + m_Key) $ 'OU'
        LOOP
      ENDIF
      
      m_Value = TRANSFORM(EVALUATE(m_FileSection + '.' + m_Key))
      m_Value = IIF(TYPE(m_FileSection + '.' + m_Key) = 'L',;
        STRTRAN(STRTRAN(m_Value, '.T.', 'yes'), '.F.', 'no'), m_Value)
      m_FileString = m_FileString +;
        SUBSTR(LOWER(m_Key), LEN(m_Prefix) + 1) + ' = ' + m_Value + CHR(13) + CHR(10)
    ENDFOR
  ENDFOR
  
  * Save IniFile
  m_RetVal = 4
      
  IF !FILE(p_File) AND !QUIET;
    AND MESSAGEBOX('Die Datei ' + ALLTRIM(p_File) + ' existiert nicht! ' +;
      'Soll sie erstellt werden?', 32 + 3, 'Einstellungen speichern') # 6
    RETURN .F.
  ENDIF

  DO WHILE m_RetVal = 4
    SET SAFETY OFF
    m_Bytes = STRTOFILE(m_FileString, p_File)
    SET SAFETY ON

    IF QUIET
      EXIT
    ENDIF 

    IF m_Bytes > 0
      m_RetVal = MESSAGEBOX("Die Einstellungen wurden erfolgreich gespeichert",;
        0, "Einstellungen speichern")
    ELSE
      m_RetVal = MESSAGEBOX("Es ist ein Fehler aufgetreten! Versuchen Sie einen anderen Dateinamen",;
        5 + 16, "Einstellungen speichern")
    ENDIF
  ENDDO

  RETURN IIF(m_Bytes > 0, m_FileString, '')
ENDFUNC