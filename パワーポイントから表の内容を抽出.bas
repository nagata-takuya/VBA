Attribute VB_Name = "Module1"
Sub パワーポイントのファイルを開く()
'
' パワーポイントのファイルを開く Macro
'

 Dim Path As String
 Dim xl_wbk As Object
  
  
  Set xl_wbk = Workbooks.Add '新規ワークブック追加
 
  With Workbooks("パワーポイントの起動用マクロ.xlsm").Worksheets("Sheet1")
      For i = 1 To .Cells(Rows.Count, 1).End(xlUp).Row
      
      
        Path = .Cells(i, 1).Value

    

        
        '-------------------------------------------------
        Call 表の全文字列をExcelに出力する(Path, xl_wbk)  '対象のマクロを選択
        '-------------------------------------------------
        
       Next i
    End With
    
End Sub


Public Function 表の全文字列をExcelに出力する(Path As String, xl_wbk As Object)

  Dim ppapp As New PowerPoint.Application
  Dim sld As PowerPoint.Slide
  Dim shp As Object
  Dim r As Long  'PowerPointの表の行番号
  Dim c As Long  'PowerPointの表の列番号
  Dim xl_app As Object
  Dim xl_row As Long  'Excelの出力先行番号
  Dim bi_tyouhyono As Integer '帳票名
  Dim bi_tyouhyoid '
  
  ppapp.Visible = True
  Set ppPR = ppapp.Presentations.Open(Path)

  '出力先設定
  xl_row = Cells(Rows.Count, 1).End(xlUp).Row + 1
  

  '
  For Each sld In ppPR.Slides
    For Each shp In sld.Shapes
      If shp.HasTable Then
      With shp.Table
         If .Columns.Count >= 2 Then
            If .Cell(1, 2).Shape.TextFrame.TextRange = "出力項目" Then
              For r = 2 To .Rows.Count
                For c = 1 To .Columns.Count
                  xl_wbk.Worksheets(1).Cells(xl_row, c).Value = _
                    .Cell(r, c).Shape.TextFrame.TextRange
                Next c
                xl_row = xl_row + 1
              Next r
            End If
        End If
      End With
      End If
    Next shp
  Next sld
  
  ppapp.Quit
  Set pptapp = Nothing
        
  Set 表の全文字列をExcelに出力する = xl_wbk
  
  End Function



