'Option Explicit

Dim strItemList As String
Global isConnected As Boolean
Public conItemIntegrity As New ADODB.Connection

Private Sub Start()

    If Range("A2").Value = "" Then
        MsgBox "You must provide at least one item starting at Cell A2 before running!", vbInformation, "Missing Data"
        Exit Sub
    End If

    If Internals.conItemIntegrity.State <> adStateOpen Then
        frmLogin.Show
    End If
    
    strItemList = ""
    Call CreateItemList
        
    Dim strSQL As String
    
    
        strSQL = "SELECT DISTINCT ITGEN_GENINFO_TBL.ITGEN_ITM_NBR AS ""ITEM"", ITGEN_GENINFO_TBL.ITGEN_ORIG_ITM_NBR AS ""ORIG ITEM"" , ITGEN_GENINFO_TBL.ITGEN_ITM_DESC AS ""ITEM DESCRIPTION"",ITGEN_LBL_DESC1 AS ""DESCRIPTOR 1"", ITGEN_LBL_DESC2 AS ""DESCRIPTOR 2"", ITGEN_LBL_DESC3 AS ""DESCRIPTOR 3"", ITIAP_ITMACCS_TBL.ITIAP_ALT_KEY_TYPE AS ""REF TYPE"", ITIAP_ITMACCS_TBL.ITIAP_ALT_KEY_VALU AS ""REF VALUE"", " & _
                 "ITIAP_ITMACCS_TBL.ITIAP_PRFD_UPC_IND AS ""PRFD IND"", ITGEN_GENINFO_TBL.ITGEN_STAT AS ""ITEM STAT"", ITGEN_GENINFO_TBL.ITGEN_ITEM_TYPE AS ""ITM TYPE"", ITGEN_GENINFO_TBL.ITGEN_ITM_SUBTYP AS ""SUB TYPE"", ITGEN_GENINFO_TBL.ITGEN_DSTN_DEPT AS ""DIST DEPT"", " & _
                 "ITDPT_ITORDDPT_TBL.ITDPT_ORD_DPT AS ""ORD DEPT"", MKLPI_LUPDMITM_TBL.MKLPI_MAX_WHSL_PRC AS ""UNT COST"", MKLPI_LUPDMITM_TBL.MKLPI_PRM_RTL_PRC AS ""UNT RTL"", (CAST(ITUPR_ITUPRSZS_TBL.ITUPR_UPR_UNTS AS CHAR(6)) || SPACE(1) || CAST(ITUPR_ITUPRSZS_TBL.ITUPR_UPR_UOM_CDE AS CHAR(2))) AS ""SIZE"", ITGEN_GENINFO_TBL.ITGEN_CURR_ITM_IND AS ""CURR ITM IND"",  " & _
                 "ITGEN_GENINFO_TBL.ITGEN_BUYG_UNTS AS ""BUY UNTS"", MKLPI_LUPDMITM_TBL.MKLPI_CATG AS ""CATG"", MKLPI_LUPDMITM_TBL.MKLPI_CATG_DESC AS ""CATG DESCRIPTION"", MKLPI_LUPDMITM_TBL.MKLPI_CLS AS ""CLASS"", MKLPI_LUPDMITM_TBL.MKLPI_CLS_DESC AS ""CLASS DESCRIPTION"", MKLPI_LUPDMITM_TBL.MKLPI_SUBCLASS AS ""SUBCL"", MKLPI_LUPDMITM_TBL.MKLPI_SUBCL_DESC AS ""SUBCL DESCRIPTION"", " & _
                 "'' AS ""DC10"", '' AS ""DC29"", '' AS ""DC53"", '' AS ""DC55"", '' AS ""DC79"", '' AS ""DC80"", '' AS ""DC81"", '' AS ""DC88"" " & _
                 "FROM SYSADM.ITGEN_GENINFO_TBL ITGEN_GENINFO_TBL INNER JOIN SYSADM.ITDPT_ITORDDPT_TBL ITDPT_ITORDDPT_TBL ON ITGEN_GENINFO_TBL.ITGEN_ITM_NBR=ITDPT_ITORDDPT_TBL.ITDPT_ITM_NBR LEFT OUTER JOIN SYSADM.MKLPI_LUPDMITM_TBL MKLPI_LUPDMITM_TBL ON ITGEN_GENINFO_TBL.ITGEN_ITM_NBR=MKLPI_LUPDMITM_TBL.MKLPI_ITM_NBR LEFT OUTER JOIN SYSADM.ITUPR_ITUPRSZS_TBL ITUPR_ITUPRSZS_TBL ON ITGEN_GENINFO_TBL.ITGEN_ITM_NBR=ITUPR_ITUPRSZS_TBL.ITUPR_ITM_NBR " & _
                 "AND ITGEN_GENINFO_TBL.ITGEN_CORP=ITUPR_ITUPRSZS_TBL.ITUPR_DSTG_SBSY LEFT OUTER JOIN SYSADM.ITIAP_ITMACCS_TBL ITIAP_ITMACCS_TBL ON ITGEN_GENINFO_TBL.ITGEN_ITM_NBR=ITIAP_ITMACCS_TBL.ITIAP_ITM_NBR AND ITGEN_GENINFO_TBL.ITGEN_CORP=ITIAP_ITMACCS_TBL.ITIAP_CORP " & _
                 "WHERE ITIAP_ITMACCS_TBL.ITIAP_ALT_KEY_TYPE = 'UPC' AND ITDPT_ITORDDPT_TBL.ITDPT_DSTG_SBSY='100' " & _
                 "AND ITIAP_ITMACCS_TBL.ITIAP_DLT_IND='N' AND ITUPR_ITUPRSZS_TBL.ITUPR_GVRNG_ATHRY='DFLT' AND ITUPR_ITUPRSZS_TBL.ITUPR_DLT_IND='N' AND ITGEN_GENINFO_TBL.ITGEN_CORP='100' AND (ITGEN_GENINFO_TBL.ITGEN_ITM_NBR IN (" & strItemList & ")) AND ITDPT_ITORDDPT_TBL.ITDPT_DLT_IND='N' " & _
                 "ORDER BY MKLPI_LUPDMITM_TBL.MKLPI_CATG_DESC, MKLPI_LUPDMITM_TBL.MKLPI_CLS_DESC, MKLPI_LUPDMITM_TBL.MKLPI_SUBCL_DESC, ITGEN_GENINFO_TBL.ITGEN_ITM_DESC;"
                

    On Error GoTo Error_Handler
    
    Dim rstItemIntegrity As New ADODB.Recordset
        rstItemIntegrity.Open strSQL, conItemIntegrity
    
    If Not rstItemIntegrity.EOF Then
    
    Application.ScreenUpdating = False
    
        Dim intHeading As Integer
            intHeading = 1
        Dim objField As ADODB.Field
        
        For Each objField In rstItemIntegrity.Fields
            ActiveWorkbook.Sheets("RESULTS").Cells(1, intHeading).Value = objField.Name
            intHeading = intHeading + 1
        Next objField
        
        ActiveWorkbook.Sheets("RESULTS").Range("A2").CopyFromRecordset rstItemIntegrity
           
    Call UpdateWhseFields
    
    ActiveWorkbook.Sheets("RESULTS").Activate
    
    Call Document_Cosmetics
        
    Else
        MsgBox "No Data Found based on item(s) provided", vbInformation, "No Data Found"
        Exit Sub
    End If

Error_Exit:
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
Error_Handler:
    MsgBox Err.Description, vbCritical, "Oops!"
    GoTo Error_Exit

End Sub

Private Sub CreateItemList()
    
    Dim lngRows As Long
        lngRows = 2
    
    Dim wsItems As Worksheet
    Set wsItems = ActiveWorkbook.Sheets("ITEMS")
        
    Do While wsItems.Cells(lngRows, 1).Value <> ""
        
        strItemList = strItemList & "'" & Format(wsItems.Cells(lngRows, 1).Value, "0000000")
            If wsItems.Cells(lngRows, 1).Offset(1, 0).Value = "" Then
                strItemList = strItemList & "'"
            Else
                strItemList = strItemList & "',"
            End If
        lngRows = lngRows + 1
    Loop
        
End Sub


Private Sub Document_Cosmetics()

    'The following blocks determine which columns to format and how to justify(position)
        'Left justified columns
        Range("A:E,Q:V").HorizontalAlignment = xlLeft
        'Right justified columns
        Range("L:N").HorizontalAlignment = xlRight
        'Center justified columns
        Range("F:K,O:P,W:AJ").HorizontalAlignment = xlCenter
    'The following formats the first row to look more appropriate to a heading
        Range("A1:AJ1").Style = "Heading 3"
        
    'The following is to autofit the data to their columns
        Cells.CurrentRegion.Columns.AutoFit
    
    'The following is to set up the data for printing
        With ActiveSheet
            With .PageSetup
                .PrintArea = Cells.CurrentRegion.Address
                .PrintTitleRows = "$1:$1"
                .PrintGridlines = True
                .Zoom = False
                .FitToPagesWide = 1
                .Orientation = xlLandscape
                .PaperSize = xlPaperLegal
                .TopMargin = "0.5"
                .BottomMargin = "0.5"
                .LeftMargin = "0.5"
                .RightMargin = "0.5"
            End With
        End With

End Sub

Private Sub ClearSheet()
    
    Application.ScreenUpdating = False
    
    ActiveWorkbook.Sheets("RESULTS").Cells.Clear
    ActiveWorkbook.Sheets("ITEMS").Range("A2:A10000").Clear
    ActiveWorkbook.Sheets("ITEMS").Activate
    ActiveSheet.Range("A2").Select
    
    Application.ScreenUpdating = True

End Sub

Private Sub UpdateWhseFields()
    
    Dim lngRows As Long
        lngRow = 2
    Dim strWhse As String
    
    Do While ActiveWorkbook.Sheets("RESULTS").Cells(lngRow, 1).Value <> ""
        
        Dim strItem As String
            strItem = ActiveWorkbook.Sheets("RESULTS").Cells(lngRow, 1).Value
            
            For i = 0 To 7 Step 1
            
            Select Case i
                Case 0:
                    strWhse = "00010"
                    ActiveWorkbook.Sheets("RESULTS").Cells(lngRow, i + 26) = IsWhseItem_Active(strItem, strWhse)
                Case 1:
                    strWhse = "00029"
                    ActiveWorkbook.Sheets("RESULTS").Cells(lngRow, i + 26) = IsWhseItem_Active(strItem, strWhse)
                Case 2:
                    strWhse = "00053"
                    ActiveWorkbook.Sheets("RESULTS").Cells(lngRow, i + 26) = IsWhseItem_Active(strItem, strWhse)
                Case 3:
                    strWhse = "00055"
                    ActiveWorkbook.Sheets("RESULTS").Cells(lngRow, i + 26) = IsWhseItem_Active(strItem, strWhse)
                Case 4:
                    strWhse = "00079"
                    ActiveWorkbook.Sheets("RESULTS").Cells(lngRow, i + 26) = IsWhseItem_Active(strItem, strWhse)
                Case 5:
                    strWhse = "00080"
                    ActiveWorkbook.Sheets("RESULTS").Cells(lngRow, i + 26) = IsWhseItem_Active(strItem, strWhse)
                Case 6:
                    strWhse = "00081"
                    ActiveWorkbook.Sheets("RESULTS").Cells(lngRow, i + 26) = IsWhseItem_Active(strItem, strWhse)
                Case 7:
                    strWhse = "00088"
                    ActiveWorkbook.Sheets("RESULTS").Cells(lngRow, i + 26) = IsWhseItem_Active(strItem, strWhse)
                Case 8:
                    strWhse = "00010"
                    ActiveWorkbook.Sheets("RESULTS").Cells(lngRow, i + 26) = IsWhseItem_Active(strItem, strWhse)
                    
            End Select
                
           Next i
           
           lngRow = lngRow + 1
    Loop

End Sub

Public Function IsWhseItem_Active(ByRef strItem As String, ByRef strWhse As String)

    Dim rstWarehouseActive As New ADODB.Recordset
    Dim strSQL As String
        
        strSQL = "SELECT COALESCE(ITWHS_STAT,'NONE') " & _
                 "FROM SYSADM.ITWHS_ITMWHS_TBL " & _
                 "WHERE ITWHS_CORP = '100' AND " & _
                 "ITWHS_ITM_NBR = '" & strItem & "' " & _
                 "AND ITWHS_DSTN_CNTR_ID = '" & strWhse & "';"
                 
        rstWarehouseActive.Open strSQL, conItemIntegrity, , adLockReadOnly
        
        If Not rstWarehouseActive.EOF Then
            IsWhseItem_Active = rstWarehouseActive.Fields(0).Value
        Else
            IsWhseItem_Active = "N/A"
        End If
        
        rstWarehouseActive.Close

End Function
