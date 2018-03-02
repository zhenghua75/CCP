Attribute VB_Name = "modRF"
Option Explicit
Global icdev As Long
Global st As Integer


'comm function
Declare Function rf_config Lib "mwrf32.dll" (ByVal icdev%, ByVal mode%, ByVal baud%) As Integer
Declare Function rf_init Lib "mwrf32.dll" (ByVal port%, ByVal baud As Long) As Long
Declare Function rf_exit Lib "mwrf32.dll" (ByVal icdev As Long) As Integer
Declare Function rf_request Lib "mwrf32.dll" (ByVal icdev As Long, ByVal mode%, atr_type%) As Integer
Declare Function rf_anticoll Lib "mwrf32.dll" (ByVal icdev As Long, ByVal Bcnt%, Snr As Long) As Integer
Declare Function rf_select Lib "mwrf32.dll" (ByVal icdev%, ByVal Snr As Long, Size%) As Integer
Declare Function rf_card Lib "mwrf32.dll" (ByVal icdev As Long, ByVal mode%, Snr As Long) As Integer
Declare Function rf_load_key Lib "mwrf32.dll" (ByVal icdev As Long, ByVal mode%, ByVal secnr%, ByRef nkey As Byte) As Integer
Declare Function rf_load_key_hex Lib "mwrf32.dll" (ByVal icdev As Long, ByVal mode%, ByVal secnr%, ByVal nkey As String) As Integer
Declare Function rf_authentication Lib "mwrf32.dll" (ByVal icdev As Long, ByVal mode%, ByVal scenr%) As Integer
Declare Function rf_read Lib "mwrf32.dll" (ByVal icdev As Long, ByVal Adr%, ByVal sdata$) As Integer
Declare Function rf_read_hex Lib "mwrf32.dll" (ByVal icdev As Long, ByVal Adr%, ByVal sdata$) As Integer
Declare Function rf_write Lib "mwrf32.dll" (ByVal icdev As Long, ByVal Adr%, ByVal sdata$) As Integer
Declare Function rf_write_hex Lib "mwrf32.dll" (ByVal icdev As Long, ByVal Adr%, ByVal sdata$) As Integer
Declare Function rf_check_write Lib "mwrf32.dll" (ByVal icdev As Long, ByVal Snr As Long, ByVal mode%, ByVal Adr%, ByVal sdata$) As Integer
Declare Function rf_check_writehex Lib "mwrf32.dll" (ByVal icdev As Long, ByVal Snr As Long, ByVal mode%, ByVal Adr%, ByVal sdata$) As Integer
'
Declare Function rf_initval Lib "mwrf32.dll" (ByVal icdev As Long, ByVal Adr%, ByVal value As Long) As Integer
Declare Function rf_readval Lib "mwrf32.dll" (ByVal icdev As Long, ByVal Adr%, value As Long) As Integer
Declare Function rf_increment Lib "mwrf32.dll" (ByVal icdev As Long, ByVal Adr%, ByVal value As Long) As Integer
Declare Function rf_decrement Lib "mwrf32.dll" (ByVal icdev As Long, ByVal Adr%, ByVal value As Long) As Integer
Declare Function rf_restore Lib "mwrf32.dll" (ByVal icdev As Long, ByVal Adr%) As Integer
Declare Function rf_transfer Lib "mwrf32.dll" (ByVal icdev As Long, ByVal Adr%) As Integer
Declare Function rf_halt Lib "mwrf32.dll" (ByVal icdev As Long) As Integer
 
'
Declare Function rf_HL_increment Lib "mwrf32.dll" (ByVal icdev As Long, ByVal mode%, ByVal secnr%, ByVal value As Long, Snr As Long, svalue As Long, ssnr As Long) As Integer
Declare Function rf_HL_decrement Lib "mwrf32.dll" (ByVal icdev As Long, ByVal mode%, ByVal secnr%, ByVal value As Long, Snr As Long, svalue As Long, ssnr As Long) As Integer
Declare Function rf_HL_write Lib "mwrf32.dll" (ByVal icdev As Long, ByVal mode%, ByVal Adr%, Snr As Long, ByVal sdata$) As Integer
Declare Function rf_HL_read Lib "mwrf32.dll" (ByVal icdev As Long, ByVal mode%, ByVal Adr%, Snr As Long, ByVal sdata$, ssnr As Long) As Integer
Declare Function rf_HL_writehex Lib "mwrf32.dll" (ByVal icdev As Long, ByVal mode%, ByVal Adr%, Snr As Long, ByVal sdata$) As Integer
Declare Function rf_HL_readhex Lib "mwrf32.dll" (ByVal icdev As Long, ByVal mode%, ByVal Adr%, Snr As Long, ByVal sdata$, Rsnr As Long) As Integer
Declare Function rf_HL_initval Lib "mwrf32.dll" (ByVal icdev As Long, ByVal mode%, ByVal secnr%, ByVal value As Long, Snr As Long) As Integer


'ML card function
Declare Function rf_initval_ml Lib "mwrf32.dll" (ByVal icdev As Long, ByVal value%) As Integer
Declare Function rf_readval_ml Lib "mwrf32.dll" (ByVal icdev As Long, value%) As Integer
Declare Function rf_decrement_ml Lib "mwrf32.dll" (ByVal icdev As Long, ByVal value%) As Integer
Declare Function rf_decrement_transfer Lib "mwrf32.dll" (ByVal icdev As Long, ByVal Adr%, ByVal value%) As Integer
Declare Function rf_authentication_2 Lib "mwrf32.dll" (ByVal icdev As Long, ByVal mode%, ByVal keynr%, ByVal Adr%) As Integer
'device fuction
Declare Function rf_reset Lib "mwrf32.dll" (ByVal icdev As Long, ByVal msec%) As Integer
Declare Function rf_get_status Lib "mwrf32.dll" (ByVal icdev As Long, ByVal status$) As Integer
Declare Function rf_encrypt Lib "mwrf32.dll" (ByVal key As String, ByVal ptrsource As String, ByVal msglen%, ByVal ptrdest$) As Integer
Declare Function rf_decrypt Lib "mwrf32.dll" (ByVal key As String, ByVal ptrsource As String, ByVal msglen%, ByVal ptrdest$) As Integer
Declare Function lib_ver Lib "mwrf32.dll" (ByVal str_ver As String) As Integer
Declare Function rf_beep Lib "mwrf32.dll" (ByVal icdev As Long, ByVal time1 As Integer) As Integer
Declare Function rf_srd_snr Lib "mwrf32.dll" (ByVal icdev As Long, ByVal offset%, ByVal rec_buffer$) As Integer
Declare Function rf_srd_eeprom Lib "mwrf32.dll" (ByVal icdev As Long, ByVal offset%, ByVal lenth%, ByVal rec_buffer$) As Integer
Declare Function rf_swr_eeprom Lib "mwrf32.dll" (ByVal icdev As Long, ByVal offset%, ByVal lenth%, ByVal send_buffer$) As Integer
Declare Function rf_changeb3 Lib "mwrf32.dll" (ByVal Adr As Long, ByVal secer As Integer, ByRef KeyA As Byte, ByVal B0 As Integer, ByVal B1 As Integer, ByVal B2 As Integer, ByVal B3 As Integer, ByVal Bk As Integer, ByRef KeyB As Byte) As Integer
Declare Function rf_disp8 Lib "mwrf32.dll" (ByVal icdev As Long, ByVal pt_mode As Integer, ByRef digit As Byte) As Integer
Declare Function rf_gettimehex Lib "mwrf32.dll" (ByVal icdev As Long, ByVal time$) As Integer
Declare Function rf_settimehex Lib "mwrf32.dll" (ByVal icdev As Long, ByVal time$) As Integer
Declare Function rf_setbright Lib "mwrf32.dll" (ByVal icdev As Long, ByVal value%) As Integer
Declare Function rf_ctl_mode Lib "mwrf32.dll" (ByVal icdev As Long, ByVal mode%) As Integer
Declare Function rf_disp_mode Lib "mwrf32.dll" (ByVal icdev As Long, ByVal mode%) As Integer
Declare Function rf_comm_check Lib "mwrf32.dll" (ByVal icdev As Long, ByVal mode%) As Integer
Declare Function set_host_check Lib "mwrf32.dll" (ByVal mode%) As Integer
Declare Function set_host_485 Lib "mwrf32.dll" (ByVal mode%) As Integer

Sub quit()
    If icdev > 0 Then
       st = rf_reset(icdev, 10)
       st = rf_exit(icdev)
       icdev = -1
    End If
End Sub





