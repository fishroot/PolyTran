LPARAMETERS p_Channel, p_Action, p_Item, p_Data, p_Format, p_Transaction

p_Data = CHRTRAN(p_Data, CHR(0), '')
DO CASE
  CASE p_Item # 'TRANSLATION'
  CASE EMPTY(p_Data)
    _Screen.Translator.mm_Request(p_Item)
  OTHERWISE
    _Screen.Translator.mp_Translation = p_Data
ENDCASE