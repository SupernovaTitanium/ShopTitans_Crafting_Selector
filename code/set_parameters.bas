



Public common_resource_num As Integer
Public uncommon_resource_num As Integer
Public rare_resource_num As Integer
Public worker_num As Integer
Public folder_Path As String
Public enchant_num As Integer
Public weapons_num As Integer
Public armor_num As Integer
Public acessories_num As Integer

Sub set_parameters()
  ' Declaration
  Dim parameters_start_column As Integer
  Dim parameters_start_row As Integer
  Dim i As Integer
  Dim j As Integer
  Dim temp_pic As Object
  Dim folder_Path As String
  Dim resource_size As Integer
  Dim resource_center_offset_left As Double
  Dim resource_center_offset_top As Double
  Dim asc_size As Integer
  Dim asc_offset_left As Double
  Dim asc_offset_top As Double
  Dim current_sheet As Worksheet
  Dim Buff_offset_left As Double
  Dim Buff_offset_top As Double
  Dim Buff_size As Integer
  Dim event_width As Integer
  Dim event_height As Integer
  Dim event_offset_left As Double
  Dim event_offset_top As Double
  Dim quality_size As Integer
  Dim quality_offset_left As Double
  Dim quality_offset_top As Double
  ' Initialization
  'Set current_sheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
  Set current_sheet = ActiveWorkbook.Sheets(1)
  folder_Path = Application.ActiveWorkbook.Path
  common_resource_num = 4
  uncommon_resource_num = 4
  rare_resource_num = 2
  worker_num = 10
  enchant_num = 2
  weapons_num = 10
  armor_num = 10
  acessories_num = 6
  ' Set the background color/fonts/alignments
  With current_sheet
        .Name = "Parameters"
        .Cells.Interior.Color = RGB(0, 0, 0)
        .Cells.Font.Color = RGB(255, 255, 255)
        .Cells.Font.Name = "HP Simplified Light"
        .Cells.HorizontalAlignment = xlCenter
        .Cells.VerticalAlignment = xlCenter
        'location
        parameters_start_column = 5
        parameters_start_row = 2
        'Title
        .Range(.Cells(parameters_start_row, parameters_start_column + 1+1), .Cells(parameters_start_row , parameters_start_column + 10+1)).Merge
        .Cells(parameters_start_row, parameters_start_column + 1+1).Interior.Color =  RGB(35, 35, 35)
        .Cells(parameters_start_row, parameters_start_column + 1+1).Value = "Resources, Workers, Buffs and Events"
        'Resources
        .Rows(parameters_start_row + 1).RowHeight = 40
        .Rows(parameters_start_row + 1).ColumnWidth = 15
        'levels
        .Rows(parameters_start_row + 2).RowHeight = 18
        'Base
        .Rows(parameters_start_row + 3).RowHeight = 18
        'Real
        .Rows(parameters_start_row + 4).RowHeight = 18
        'resource Image
        resource_size = 35
        resource_center_offset_left = .Rows(parameters_start_row + 1).ColumnWidth * 5.25/ 2 - resource_size / 2
        resource_center_offset_top = .Rows(parameters_start_row + 1).RowHeight / 2 - resource_size / 2
        For i = 1 To common_resource_num + uncommon_resource_num + rare_resource_num
            .Cells(parameters_start_row + 1, parameters_start_column + 1 + i).Interior.Color = RGB(55, 55, 55)
            Set temp_pic = .Shapes.AddPicture(folder_Path & "\Images\Resources\" & CStr(i) & ".png", True, True, .Cells(parameters_start_row + 1, parameters_start_column + 1 + i).Left + resource_center_offset_left, .Cells(parameters_start_row + 1, parameters_start_column + 1 + i).Top + resource_center_offset_top, resource_size, resource_size)
            .Cells(parameters_start_row + 2, parameters_start_column + 1 + i).Interior.Color = RGB(255, 0, 0)
            .Cells(parameters_start_row + 2, parameters_start_column + 1 + i).Value = 16
            .Cells(parameters_start_row + 3, parameters_start_column + 1 + i).Interior.Color = RGB(100, 100, 100)
            .Cells(parameters_start_row + 3, parameters_start_column + 1 + i).NumberFormat = "0.00" & Chr$(34) & "/min" & Chr$(34)
            Select Case i
                Case Is < 5
                    .Cells(parameters_start_row + 3, parameters_start_column + 1 + i).Formula = "=" & _
                    "IF(" & .Cells(parameters_start_row + 2, parameters_start_column + 1 + i).Address(False, False) & _
                    "<10,(" & .Cells(parameters_start_row + 2, parameters_start_column + 1 + i).Address(False, False) & _
                    "-1)*0.5+6,IF(" & .Cells(parameters_start_row + 2, parameters_start_column + 1 + i).Address(False, False) & _
                    "<18,(" & .Cells(parameters_start_row + 2, parameters_start_column + 1 + i).Address(False, False) & "-9)*0.75+10,IF(" & _
                    .Cells(parameters_start_row + 2, parameters_start_column + 1 + i).Address(False, False) & _
                    "<20,(" & .Cells(parameters_start_row + 2, parameters_start_column + 1 + i).Address(False, False) & "-18)*2+17,22)))"
                Case Is < 9
                    .Cells(parameters_start_row + 3, parameters_start_column + 1 + i).Formula = "=" & "IF(" & .Cells(parameters_start_row + 2, parameters_start_column + 1 + i).Address(False, False) & _
                    "<10,(" & .Cells(parameters_start_row + 2, parameters_start_column + 1 + i).Address(False, False) & "-1)*0.2+0.7,IF(" & _
                    .Cells(parameters_start_row + 2, parameters_start_column + 1 + i).Address(False, False) & "<18,(" & .Cells(parameters_start_row + 2, parameters_start_column + 1 + i).Address(False, False) & "-9)*0.3+2.3,IF(" & _
                    .Cells(parameters_start_row + 2, parameters_start_column + 1 + i).Address(False, False) & "<20,(" & .Cells(parameters_start_row + 2, parameters_start_column + 1 + i).Address(False, False) & "-18)*0.8+5.2,7)))"
                Case Else
                    .Cells(parameters_start_row + 3, parameters_start_column + 1 + i).Formula = "=" & "IF(" & .Cells(parameters_start_row + 2, parameters_start_column + 1 + i).Address(False, False) & _
                    "<18,(" & .Cells(parameters_start_row + 2, parameters_start_column + 1 + i).Address(False, False) & "-1)*0.05+0.1,IF(" & _
                    .Cells(parameters_start_row + 2, parameters_start_column + 1 + i).Address(False, False) & "<20,(" & .Cells(parameters_start_row + 2, parameters_start_column + 1 + i).Address(False, False) & "-18)*0.1+1,1.25))"
            End Select
            .Cells(parameters_start_row + 4, parameters_start_column + 1 + i).Interior.Color = RGB(8, 46, 84)
            .Cells(parameters_start_row + 4, parameters_start_column + 1 + i).NumberFormat = "0.0000" & Chr$(34) & "/min" & Chr$(34)
            .Cells(parameters_start_row + 4, parameters_start_column + 1 + i).Formula = "=" & .Cells(parameters_start_row + 3, parameters_start_column + 1 + i).Address(False, False) & _
            "*(1+" & .Cells(parameters_start_row + 10, parameters_start_column + 3).Address(False, False) & "+" & .Cells(parameters_start_row + 9, parameters_start_column + 5).Address(False, False) _
            & "*" & .Cells(parameters_start_row + 10, parameters_start_column + 5).Address(False, False) & ")"
        Next
        'Workers
        .Rows(parameters_start_row + 5).RowHeight = 40
        .Rows(parameters_start_row + 5).ColumnWidth = 15
        'levels
        .Rows(parameters_start_row + 6).RowHeight = 18
        'Speed up
        .Rows(parameters_start_row + 7).RowHeight = 18
        'Workers Image
        worker_size_height = 35
        worker_size_width = 35
        worker_center_offset_left = .Rows(parameters_start_row + 5).ColumnWidth * 5.25/ 2 - worker_size_width / 2
        worker_center_offset_top = .Rows(parameters_start_row + 5).RowHeight / 2 - worker_size_height / 2
        For i = 1 To worker_num
            .Cells(parameters_start_row + 5, parameters_start_column + 1 + i).Interior.Color = RGB(55, 55, 55)
            Set temp_pic = .Shapes.AddPicture(folder_Path & "\Images\Workers\" & CStr(i) & ".png", True, True, .Cells(parameters_start_row + 5, parameters_start_column + 1 + i).Left + worker_center_offset_left, .Cells(parameters_start_row + 5, parameters_start_column + 1 + i).Top + worker_center_offset_top, worker_size_height, worker_size_width)
            'Workers Level
            .Cells(parameters_start_row + 6, parameters_start_column + 1 + i).Interior.Color = RGB(255, 0, 0)
            .Cells(parameters_start_row + 6, parameters_start_column + 1 + i).Value = 33
            'Speedup
            .Cells(parameters_start_row + 7, parameters_start_column + 1 + i).Interior.Color = RGB(8, 46, 84)
            .Cells(parameters_start_row + 7, parameters_start_column + 1 + i).NumberFormat = Chr$(34) & "+" & Chr$(34) & "0.00%"
            .Cells(parameters_start_row + 7, parameters_start_column + 1 + i).Formula = "=IF(" & .Cells(parameters_start_row + 6, parameters_start_column + 1 + i).Address(False, False) & "=40,0.6," & _
            "ROUNDUP((" & .Cells(parameters_start_row + 6, parameters_start_column + 1 + i).Address(False, False) & "-1)/2,0)*0.01+ROUNDDOWN((" & _
            .Cells(parameters_start_row + 6, parameters_start_column + 1 + i).Address(False, False) & "-1)/2,0)*0.02)"
        Next i
        'Guild Buffs
        .Rows(parameters_start_row + 8).RowHeight = 40
        .Rows(parameters_start_row + 8).ColumnWidth = 15
        Buff_size = 35
        Buff_offset_left = .Rows(parameters_start_row + 8).ColumnWidth * 5.25/ 2 - Buff_size / 2
        Buff_offset_top = .Rows(parameters_start_row + 8).RowHeight / 2 - Buff_size / 2

        
        For i = 1 To 3
        .Cells(parameters_start_row + 8, parameters_start_column + 1 + i).Interior.Color = RGB(55, 55, 55)
        Set temp_pic = .Shapes.AddPicture(folder_Path & "\Images\Guild_Buffs\" & CStr(i) & ".png", True, True, .Cells(parameters_start_row + 8, parameters_start_column + 1 + i).Left + Buff_offset_left, .Cells(parameters_start_row + 8, parameters_start_column + 1 + i).Top + Buff_offset_top, Buff_size, Buff_size)
        Next i
        'level
        .Rows(parameters_start_row + 9).RowHeight = 18
        .Cells(parameters_start_row + 9, parameters_start_column + 2).Value = 5
        .Cells(parameters_start_row + 9, parameters_start_column + 2).Interior.Color = RGB(255, 0, 0)
        .Cells(parameters_start_row + 9, parameters_start_column + 3).Value = 5
        .Cells(parameters_start_row + 9, parameters_start_column + 3).Interior.Color = RGB(255, 0, 0)
        .Cells(parameters_start_row + 9, parameters_start_column + 4).Value = 5
        .Cells(parameters_start_row + 9, parameters_start_column + 4).Interior.Color = RGB(255, 0, 0)
        'Boost
        .Rows(parameters_start_row + 10).RowHeight = 18
        .Cells(parameters_start_row + 10, parameters_start_column + 2).NumberFormat = Chr$(34) & "+" & Chr$(34) & "0.00%"
        .Cells(parameters_start_row + 10, parameters_start_column + 3).NumberFormat = Chr$(34) & "+" & Chr$(34) & "0.00%"
        .Cells(parameters_start_row + 10, parameters_start_column + 4).NumberFormat = Chr$(34) & "+" & Chr$(34) & "0.00%"
        .Cells(parameters_start_row + 10, parameters_start_column + 2).Formula = "=If(" & .Cells(parameters_start_row + 9, parameters_start_column + 2).Address(False, False) & ">0," & .Cells(parameters_start_row + 9, parameters_start_column + 2).Address(False, False) & "*0.05+0.05,0)"
        .Cells(parameters_start_row + 10, parameters_start_column + 3).Formula = "=" & .Cells(parameters_start_row + 9, parameters_start_column + 3).Address(False, False) & "*0.1"
        .Cells(parameters_start_row + 10, parameters_start_column + 4).Formula = "=" & .Cells(parameters_start_row + 9, parameters_start_column + 4).Address(False, False) & "*0.05"
        .Cells(parameters_start_row + 10, parameters_start_column + 2).Interior.Color = RGB(8, 46, 84)
        .Cells(parameters_start_row + 10, parameters_start_column + 3).Interior.Color = RGB(8, 46, 84)
        .Cells(parameters_start_row + 10, parameters_start_column + 4).Interior.Color = RGB(8, 46, 84)
        

    'Events
        event_height = 38
        event_width = 43
        event_offset_left = .Rows(parameters_start_row + 8).ColumnWidth * 5.25/ 2 - event_width / 2
        event_offset_top = .Rows(parameters_start_row + 8).RowHeight / 2 - event_height / 2
        For i = 5 To 11
            Select Case i
                Case 5
                    Set temp_pic = .Shapes.AddPicture(folder_Path & "\Images\Events\" & CStr(i) - 4 & ".png", True, True, .Cells(parameters_start_row + 8, parameters_start_column + i).Left _
                    + event_offset_left, .Cells(parameters_start_row + 8, parameters_start_column + 1 + i).Top + event_offset_top, event_width, event_height)
                    .Cells(parameters_start_row + 8, parameters_start_column + i).Interior.Color = RGB(55, 55, 55)
                    .Cells(parameters_start_row + 9, parameters_start_column + i).Value = 0
                    .Cells(parameters_start_row + 9, parameters_start_column + i).Interior.Color = RGB(255, 0, 0)
                    .Cells(parameters_start_row + 10, parameters_start_column + i).Value = 0.25
                    .Cells(parameters_start_row + 10, parameters_start_column + i).Interior.Color = RGB(8, 46, 84)
                    .Cells(parameters_start_row + 10, parameters_start_column + i).NumberFormat = Chr$(34) & "+" & Chr$(34) & "0.00%"
                Case 6
                    Set temp_pic = .Shapes.AddPicture(folder_Path & "\Images\Events\" & CStr(i) - 4 & ".png", True, True, .Cells(parameters_start_row + 8, parameters_start_column + i).Left _
                    + event_offset_left, .Cells(parameters_start_row + 8, parameters_start_column + 1 + i).Top + event_offset_top, event_width, event_height)
                    .Cells(parameters_start_row + 8, parameters_start_column + i).Interior.Color = RGB(55, 55, 55)
                    .Cells(parameters_start_row + 9, parameters_start_column + i).Value = 0
                    .Cells(parameters_start_row + 9, parameters_start_column + i).Interior.Color = RGB(255, 0, 0)
                    .Cells(parameters_start_row + 10, parameters_start_column + i).Value = 0.25
                    .Cells(parameters_start_row + 10, parameters_start_column + i).Interior.Color = RGB(8, 46, 84)
                    .Cells(parameters_start_row + 10, parameters_start_column + i).NumberFormat = Chr$(34) & "+" & Chr$(34) & "0.00%"
                Case 7
                    Set temp_pic = .Shapes.AddPicture(folder_Path & "\Images\Events\" & CStr(i) - 4 & ".png", True, True, .Cells(parameters_start_row + 8, parameters_start_column + i).Left _
                    + event_offset_left, .Cells(parameters_start_row + 8, parameters_start_column + 1 + i).Top + event_offset_top, event_width, event_height)
                    .Cells(parameters_start_row + 8, parameters_start_column + i).Interior.Color = RGB(55, 55, 55)
                    .Cells(parameters_start_row + 9, parameters_start_column + i).Value = 0
                    .Cells(parameters_start_row + 9, parameters_start_column + i).Interior.Color = RGB(255, 0, 0)
                    .Cells(parameters_start_row + 10, parameters_start_column + i).Value = 0.25
                    .Cells(parameters_start_row + 10, parameters_start_column + i).Interior.Color = RGB(8, 46, 84)
                    .Cells(parameters_start_row + 10, parameters_start_column + i).NumberFormat = Chr$(34) & "+" & Chr$(34) & "0.00%"
                Case 8
                    Set temp_pic = .Shapes.AddPicture(folder_Path & "\Images\Events\" & CStr(i) - 4 & ".png", True, True, .Cells(parameters_start_row + 8, parameters_start_column + i).Left _
                    + event_offset_left, .Cells(parameters_start_row + 8, parameters_start_column + 1 + i).Top + event_offset_top, event_width, event_height)
                    .Cells(parameters_start_row + 8, parameters_start_column + i).Interior.Color = RGB(55, 55, 55)
                    .Cells(parameters_start_row + 9, parameters_start_column + i).Value = 0
                    .Cells(parameters_start_row + 9, parameters_start_column + i).Interior.Color = RGB(255, 0, 0)
                    .Cells(parameters_start_row + 10, parameters_start_column + i).Value = 0.5
                    .Cells(parameters_start_row + 10, parameters_start_column + i).Interior.Color = RGB(8, 46, 84)
                    .Cells(parameters_start_row + 10, parameters_start_column + i).NumberFormat = Chr$(34) & "+" & Chr$(34) & "0.00%"
                Case 9
                    Set temp_pic = .Shapes.AddPicture(folder_Path & "\Images\Events\" & CStr(i) - 4 & ".png", True, True, .Cells(parameters_start_row + 8, parameters_start_column + i).Left _
                    + event_offset_left, .Cells(parameters_start_row + 8, parameters_start_column + 1 + i).Top + event_offset_top, event_width, event_height)
                    .Cells(parameters_start_row + 8, parameters_start_column + i).Interior.Color = RGB(55, 55, 55)
                    .Cells(parameters_start_row + 9, parameters_start_column + i).Value = 0
                    .Cells(parameters_start_row + 9, parameters_start_column + i).Interior.Color = RGB(255, 0, 0)
                    .Cells(parameters_start_row + 10, parameters_start_column + i).Value = 0.25
                    .Cells(parameters_start_row + 10, parameters_start_column + i).Interior.Color = RGB(8, 46, 84)
                    .Cells(parameters_start_row + 10, parameters_start_column + i).NumberFormat = Chr$(34) & "+" & Chr$(34) & "0.00%"
                Case 10
                    Set temp_pic = .Shapes.AddPicture(folder_Path & "\Images\Events\" & CStr(i) - 4 & ".png", True, True, .Cells(parameters_start_row + 8, parameters_start_column + i).Left _
                    + event_offset_left, .Cells(parameters_start_row + 8, parameters_start_column + 1 + i).Top + event_offset_top, event_width, event_height)
                    .Cells(parameters_start_row + 8, parameters_start_column + i).Interior.Color = RGB(55, 55, 55)
                    .Cells(parameters_start_row + 9, parameters_start_column + i).Value = 0
                    .Cells(parameters_start_row + 9, parameters_start_column + i).Interior.Color = RGB(255, 0, 0)
                    .Cells(parameters_start_row + 10, parameters_start_column + i).Value = -0.25
                    .Cells(parameters_start_row + 10, parameters_start_column + i).Interior.Color = RGB(8, 46, 84)
                    .Cells(parameters_start_row + 10, parameters_start_column + i).NumberFormat = "0.00%"
                Case Else
                    Set temp_pic = .Shapes.AddPicture(folder_Path & "\Images\Events\" & CStr(i) - 4 & ".png", True, True, .Cells(parameters_start_row + 8, parameters_start_column + i).Left _
                    + event_offset_left, .Cells(parameters_start_row + 8, parameters_start_column + 1 + i).Top + event_offset_top, event_width, event_height)
                    .Cells(parameters_start_row + 8, parameters_start_column + i).Interior.Color = RGB(55, 55, 55)
                    .Cells(parameters_start_row + 9, parameters_start_column + i).Value = 0
                    .Cells(parameters_start_row + 9, parameters_start_column + i).Interior.Color = RGB(255, 0, 0)
                    .Cells(parameters_start_row + 10, parameters_start_column + i).Value = 0.25
                    .Cells(parameters_start_row + 10, parameters_start_column + i).Interior.Color = RGB(8, 46, 84)
                    .Cells(parameters_start_row + 10, parameters_start_column + i).NumberFormat = Chr$(34) & "+" & Chr$(34) & "0.00%"
            End Select
        Next i
        
        
        'Relocate
        parameters_start_column = 1
        parameters_start_row = 13
        'merge Ascention tress
        For i = 1 To 4 * 5
            .Cells(parameters_start_row + 1, parameters_start_column + i).Interior.Color = RGB(35, 35, 35)
        Next i
        .Range(.Cells(parameters_start_row + 1, parameters_start_column + 1), .Cells(parameters_start_row + 1, parameters_start_column + 20)).Merge
        .Cells(parameters_start_row + 1, parameters_start_column + 1).Value = "Ascension Trees"
        'Level
        .Cells(parameters_start_row + 2, parameters_start_column + 1).Value = "Level"
        .Cells(parameters_start_row + 2, parameters_start_column + 1).Interior.Color = RGB(65, 65, 65)
        .Cells(parameters_start_row + 2, parameters_start_column + 2).Value = "Novice"
        .Cells(parameters_start_row + 2, parameters_start_column + 2).Interior.Color = RGB(80, 80, 80)
        .Cells(parameters_start_row + 2, parameters_start_column + 3).Value = "Initiate"
        .Cells(parameters_start_row + 2, parameters_start_column + 3).Interior.Color = RGB(80, 80, 80)
        .Cells(parameters_start_row + 2, parameters_start_column + 4).Value = "Specialist"
        .Cells(parameters_start_row + 2, parameters_start_column + 4).Interior.Color = RGB(80, 80, 80)
        .Cells(parameters_start_row + 2, parameters_start_column + 5).Value = "Expert"
        .Cells(parameters_start_row + 2, parameters_start_column + 5).Interior.Color = RGB(80, 80, 80)
        'Boost
        .Range(.Cells(parameters_start_row + 3, parameters_start_column + 1), .Cells(parameters_start_row + 4, parameters_start_column + 1)).Merge
        .Cells(parameters_start_row + 3, parameters_start_column + 1).Value = "Talent"
        .Cells(parameters_start_row + 3, parameters_start_column + 1).Interior.Color = RGB(75, 75, 75)
        .Cells(parameters_start_row + 3, parameters_start_column + 2).Value = "XP"
        .Cells(parameters_start_row + 3, parameters_start_column + 2).Interior.Color = RGB(90, 90, 90)
        .Cells(parameters_start_row + 3, parameters_start_column + 3).Value = "Multicraft"
        .Cells(parameters_start_row + 3, parameters_start_column + 3).Interior.Color = RGB(90, 90, 90)
        .Cells(parameters_start_row + 3, parameters_start_column + 4).Value = "Multicraft"
        .Cells(parameters_start_row + 3, parameters_start_column + 4).Interior.Color = RGB(90, 90, 90)
        .Cells(parameters_start_row + 3, parameters_start_column + 5).Value = "Quality"
        .Cells(parameters_start_row + 3, parameters_start_column + 5).Interior.Color = RGB(90, 90, 90)
        'boost status
        .Cells(parameters_start_row + 4, parameters_start_column + 2).Value = 0.25
        .Cells(parameters_start_row + 4, parameters_start_column + 2).Interior.Color = RGB(100, 100, 100)
        .Cells(parameters_start_row + 4, parameters_start_column + 2).NumberFormat = Chr$(34) & "+" & Chr$(34) & "0.00%"
        .Cells(parameters_start_row + 4, parameters_start_column + 3).Value = 0.05
        .Cells(parameters_start_row + 4, parameters_start_column + 3).Interior.Color = RGB(100, 100, 100)
        .Cells(parameters_start_row + 4, parameters_start_column + 3).NumberFormat = Chr$(34) & "+" & Chr$(34) & "0.00%"
        .Cells(parameters_start_row + 4, parameters_start_column + 4).Value = 0.05
        .Cells(parameters_start_row + 4, parameters_start_column + 4).Interior.Color = RGB(100, 100, 100)
        .Cells(parameters_start_row + 4, parameters_start_column + 4).NumberFormat = Chr$(34) & "+" & Chr$(34) & "0.00%"
        .Cells(parameters_start_row + 4, parameters_start_column + 5).Value = 0.5
        .Cells(parameters_start_row + 4, parameters_start_column + 5).Interior.Color = RGB(100, 100, 100)
        .Cells(parameters_start_row + 4, parameters_start_column + 5).NumberFormat = Chr$(34) & "+" & Chr$(34) & "0.00%"

        'Enchantments Image
        asc_size = 25
        parameters_start_row = parameters_start_row + 4
        For i = 1 To enchant_num
            .Rows(parameters_start_row + i).RowHeight = asc_size
            .Cells(parameters_start_row + i, parameters_start_column + 1).Interior.Color = RGB(8, 46, 84)
            asc_offset_left = .Rows(parameters_start_row + i).ColumnWidth * 5.25/ 2 - asc_size / 2
            asc_offset_top = .Rows(parameters_start_row + i).RowHeight / 2 - asc_size / 2
        Set temp_pic = .Shapes.AddPicture(folder_Path & "\Images\Ascensions\Enchantments\" & CStr(i) & ".png", True, True, .Cells(parameters_start_row + i, parameters_start_column + 1).Left _
        + asc_offset_left, .Cells(parameters_start_row + i, parameters_start_column + 1).Top + asc_offset_top, asc_size, asc_size)
        For j = 1 To 4
            .Cells(parameters_start_row + i, parameters_start_column + 1 + j).Value = 0
            .Cells(parameters_start_row + i, parameters_start_column + 1 + j).Interior.Color = RGB(255, 0, 0)
        Next j
        Next i

        '''''''''''''Weapons
        parameters_start_row = 13
        parameters_start_column = 6
        'Level
        .Cells(parameters_start_row + 2, parameters_start_column + 1).Value = "Level"
        .Cells(parameters_start_row + 2, parameters_start_column + 1).Interior.Color = RGB(65, 65, 65)
        .Cells(parameters_start_row + 2, parameters_start_column + 2).Value = "Novice"
        .Cells(parameters_start_row + 2, parameters_start_column + 2).Interior.Color = RGB(80, 80, 80)
        .Cells(parameters_start_row + 2, parameters_start_column + 3).Value = "Initiate"
        .Cells(parameters_start_row + 2, parameters_start_column + 3).Interior.Color = RGB(80, 80, 80)
        .Cells(parameters_start_row + 2, parameters_start_column + 4).Value = "Specialist"
        .Cells(parameters_start_row + 2, parameters_start_column + 4).Interior.Color = RGB(80, 80, 80)
        .Cells(parameters_start_row + 2, parameters_start_column + 5).Value = "Expert"
        .Cells(parameters_start_row + 2, parameters_start_column + 5).Interior.Color = RGB(80, 80, 80)
        'Boost
        .Range(.Cells(parameters_start_row + 3, parameters_start_column + 1), .Cells(parameters_start_row + 4, parameters_start_column + 1)).Merge
        .Cells(parameters_start_row + 3, parameters_start_column + 1).Value = "Talent"
        .Cells(parameters_start_row + 3, parameters_start_column + 1).Interior.Color = RGB(75, 75, 75)
        .Cells(parameters_start_row + 3, parameters_start_column + 2).Value = "XP"
        .Cells(parameters_start_row + 3, parameters_start_column + 2).Interior.Color = RGB(90, 90, 90)
        .Cells(parameters_start_row + 3, parameters_start_column + 3).Value = "Energy"
        .Cells(parameters_start_row + 3, parameters_start_column + 3).Interior.Color = RGB(90, 90, 90)
        .Cells(parameters_start_row + 3, parameters_start_column + 4).Value = "Multicraft"
        .Cells(parameters_start_row + 3, parameters_start_column + 4).Interior.Color = RGB(90, 90, 90)
        .Cells(parameters_start_row + 3, parameters_start_column + 5).Value = "Quality"
        .Cells(parameters_start_row + 3, parameters_start_column + 5).Interior.Color = RGB(90, 90, 90)
        'boost status
        .Cells(parameters_start_row + 4, parameters_start_column + 2).Value = 0.25
        .Cells(parameters_start_row + 4, parameters_start_column + 2).Interior.Color = RGB(100, 100, 100)
        .Cells(parameters_start_row + 4, parameters_start_column + 2).NumberFormat = Chr$(34) & "+" & Chr$(34) & "0.00%"
        .Cells(parameters_start_row + 4, parameters_start_column + 3).Value = -0.1
        .Cells(parameters_start_row + 4, parameters_start_column + 3).Interior.Color = RGB(100, 100, 100)
        .Cells(parameters_start_row + 4, parameters_start_column + 3).NumberFormat = "0.00%"
        .Cells(parameters_start_row + 4, parameters_start_column + 4).Value = 0.05
        .Cells(parameters_start_row + 4, parameters_start_column + 4).Interior.Color = RGB(100, 100, 100)
        .Cells(parameters_start_row + 4, parameters_start_column + 4).NumberFormat = Chr$(34) & "+" & Chr$(34) & "0.00%"
        .Cells(parameters_start_row + 4, parameters_start_column + 5).Value = 0.5
        .Cells(parameters_start_row + 4, parameters_start_column + 5).Interior.Color = RGB(100, 100, 100)
        .Cells(parameters_start_row + 4, parameters_start_column + 5).NumberFormat = Chr$(34) & "+" & Chr$(34) & "0.00%"
        
        'Weapons Image
        asc_size = 25
        parameters_start_row = parameters_start_row + 4
        For i = 1 To weapons_num
            .Rows(parameters_start_row + i).RowHeight = asc_size
            .Cells(parameters_start_row + i, parameters_start_column + 1).Interior.Color = RGB(8, 46, 84)
            asc_offset_left = .Rows(parameters_start_row + i).ColumnWidth * 5.25/ 2 - asc_size / 2
            asc_offset_top = .Rows(parameters_start_row + i).RowHeight / 2 - asc_size / 2
        Set temp_pic = .Shapes.AddPicture(folder_Path & "\Images\Ascensions\Weapons\" & CStr(i) & ".png", True, True, .Cells(parameters_start_row + i, parameters_start_column + 1).Left _
        + asc_offset_left, .Cells(parameters_start_row + i, parameters_start_column + 1).Top + asc_offset_top, asc_size, asc_size)
        For j = 1 To 4
            .Cells(parameters_start_row + i, parameters_start_column + 1 + j).Value = 0
            .Cells(parameters_start_row + i, parameters_start_column + 1 + j).Interior.Color = RGB(255, 0, 0)
        Next j
        Next i

        
        '''''''''''''Armor
        parameters_start_row = 13
        parameters_start_column = 11
        'Level
        .Cells(parameters_start_row + 2, parameters_start_column + 1).Value = "Level"
        .Cells(parameters_start_row + 2, parameters_start_column + 1).Interior.Color = RGB(65, 65, 65)
        .Cells(parameters_start_row + 2, parameters_start_column + 2).Value = "Novice"
        .Cells(parameters_start_row + 2, parameters_start_column + 2).Interior.Color = RGB(80, 80, 80)
        .Cells(parameters_start_row + 2, parameters_start_column + 3).Value = "Initiate"
        .Cells(parameters_start_row + 2, parameters_start_column + 3).Interior.Color = RGB(80, 80, 80)
        .Cells(parameters_start_row + 2, parameters_start_column + 4).Value = "Specialist"
        .Cells(parameters_start_row + 2, parameters_start_column + 4).Interior.Color = RGB(80, 80, 80)
        .Cells(parameters_start_row + 2, parameters_start_column + 5).Value = "Expert"
        .Cells(parameters_start_row + 2, parameters_start_column + 5).Interior.Color = RGB(80, 80, 80)
        'Boost
        .Range(.Cells(parameters_start_row + 3, parameters_start_column + 1), .Cells(parameters_start_row + 4, parameters_start_column + 1)).Merge
        .Cells(parameters_start_row + 3, parameters_start_column + 1).Value = "Talent"
        .Cells(parameters_start_row + 3, parameters_start_column + 1).Interior.Color = RGB(75, 75, 75)
        .Cells(parameters_start_row + 3, parameters_start_column + 2).Value = "XP"
        .Cells(parameters_start_row + 3, parameters_start_column + 2).Interior.Color = RGB(90, 90, 90)
        .Cells(parameters_start_row + 3, parameters_start_column + 3).Value = "Energy"
        .Cells(parameters_start_row + 3, parameters_start_column + 3).Interior.Color = RGB(90, 90, 90)
        .Cells(parameters_start_row + 3, parameters_start_column + 4).Value = "Multicraft"
        .Cells(parameters_start_row + 3, parameters_start_column + 4).Interior.Color = RGB(90, 90, 90)
        .Cells(parameters_start_row + 3, parameters_start_column + 5).Value = "Quality"
        .Cells(parameters_start_row + 3, parameters_start_column + 5).Interior.Color = RGB(90, 90, 90)
        'boost status
        .Cells(parameters_start_row + 4, parameters_start_column + 2).Value = 0.25
        .Cells(parameters_start_row + 4, parameters_start_column + 2).Interior.Color = RGB(100, 100, 100)
        .Cells(parameters_start_row + 4, parameters_start_column + 2).NumberFormat = Chr$(34) & "+" & Chr$(34) & "0.00%"
        .Cells(parameters_start_row + 4, parameters_start_column + 3).Value = -0.1
        .Cells(parameters_start_row + 4, parameters_start_column + 3).Interior.Color = RGB(100, 100, 100)
        .Cells(parameters_start_row + 4, parameters_start_column + 3).NumberFormat = "0.00%"
        .Cells(parameters_start_row + 4, parameters_start_column + 4).Value = 0.05
        .Cells(parameters_start_row + 4, parameters_start_column + 4).Interior.Color = RGB(100, 100, 100)
        .Cells(parameters_start_row + 4, parameters_start_column + 4).NumberFormat = Chr$(34) & "+" & Chr$(34) & "0.00%"
        .Cells(parameters_start_row + 4, parameters_start_column + 5).Value = 0.5
        .Cells(parameters_start_row + 4, parameters_start_column + 5).Interior.Color = RGB(100, 100, 100)
        .Cells(parameters_start_row + 4, parameters_start_column + 5).NumberFormat = Chr$(34) & "+" & Chr$(34) & "0.00%"

        'Armor Image
        asc_size = 25
        parameters_start_row = parameters_start_row + 4
        For i = 1 To armor_num
            .Rows(parameters_start_row + i).RowHeight = asc_size
            .Cells(parameters_start_row + i, parameters_start_column + 1).Interior.Color = RGB(8, 46, 84)
            asc_offset_left = .Rows(parameters_start_row + i).ColumnWidth * 5.25/ 2 - asc_size / 2
            asc_offset_top = .Rows(parameters_start_row + i).RowHeight / 2 - asc_size / 2
        Set temp_pic = .Shapes.AddPicture(folder_Path & "\Images\Ascensions\Armor\" & CStr(i) & ".png", True, True, .Cells(parameters_start_row + i, parameters_start_column + 1).Left _
        + asc_offset_left, .Cells(parameters_start_row + i, parameters_start_column + 1).Top + asc_offset_top, asc_size, asc_size)
            For j = 1 To 4
                .Cells(parameters_start_row + i, parameters_start_column + 1 + j).Value = 0
                .Cells(parameters_start_row + i, parameters_start_column + 1 + j).Interior.Color = RGB(255, 0, 0)
            Next j
        Next i
        
        '''''''''''''Acessories
        parameters_start_row = 13
        parameters_start_column = 16
        'Level
        .Cells(parameters_start_row + 2, parameters_start_column + 1).Value = "Level"
        .Cells(parameters_start_row + 2, parameters_start_column + 1).Interior.Color = RGB(65, 65, 65)
        .Cells(parameters_start_row + 2, parameters_start_column + 2).Value = "Novice"
        .Cells(parameters_start_row + 2, parameters_start_column + 2).Interior.Color = RGB(80, 80, 80)
        .Cells(parameters_start_row + 2, parameters_start_column + 3).Value = "Initiate"
        .Cells(parameters_start_row + 2, parameters_start_column + 3).Interior.Color = RGB(80, 80, 80)
        .Cells(parameters_start_row + 2, parameters_start_column + 4).Value = "Specialist"
        .Cells(parameters_start_row + 2, parameters_start_column + 4).Interior.Color = RGB(80, 80, 80)
        .Cells(parameters_start_row + 2, parameters_start_column + 5).Value = "Expert"
        .Cells(parameters_start_row + 2, parameters_start_column + 5).Interior.Color = RGB(80, 80, 80)
        'Boost
        .Range(.Cells(parameters_start_row + 3, parameters_start_column + 1), .Cells(parameters_start_row + 4, parameters_start_column + 1)).Merge
        .Cells(parameters_start_row + 3, parameters_start_column + 1).Value = "Talent"
        .Cells(parameters_start_row + 3, parameters_start_column + 1).Interior.Color = RGB(75, 75, 75)
        .Cells(parameters_start_row + 3, parameters_start_column + 2).Value = "XP"
        .Cells(parameters_start_row + 3, parameters_start_column + 2).Interior.Color = RGB(90, 90, 90)
        .Cells(parameters_start_row + 3, parameters_start_column + 3).Value = "Energy"
        .Cells(parameters_start_row + 3, parameters_start_column + 3).Interior.Color = RGB(90, 90, 90)
        .Cells(parameters_start_row + 3, parameters_start_column + 4).Value = "Multicraft"
        .Cells(parameters_start_row + 3, parameters_start_column + 4).Interior.Color = RGB(90, 90, 90)
        .Cells(parameters_start_row + 3, parameters_start_column + 5).Value = "Quality"
        .Cells(parameters_start_row + 3, parameters_start_column + 5).Interior.Color = RGB(90, 90, 90)
        'boost status
        .Cells(parameters_start_row + 4, parameters_start_column + 2).Value = 0.25
        .Cells(parameters_start_row + 4, parameters_start_column + 2).Interior.Color = RGB(100, 100, 100)
        .Cells(parameters_start_row + 4, parameters_start_column + 2).NumberFormat = Chr$(34) & "+" & Chr$(34) & "0.00%"
        .Cells(parameters_start_row + 4, parameters_start_column + 3).Value = -0.1
        .Cells(parameters_start_row + 4, parameters_start_column + 3).Interior.Color = RGB(100, 100, 100)
        .Cells(parameters_start_row + 4, parameters_start_column + 3).NumberFormat = "0.00%"
        .Cells(parameters_start_row + 4, parameters_start_column + 4).Value = 0.05
        .Cells(parameters_start_row + 4, parameters_start_column + 4).Interior.Color = RGB(100, 100, 100)
        .Cells(parameters_start_row + 4, parameters_start_column + 4).NumberFormat = Chr$(34) & "+" & Chr$(34) & "0.00%"
        .Cells(parameters_start_row + 4, parameters_start_column + 5).Value = 0.5
        .Cells(parameters_start_row + 4, parameters_start_column + 5).Interior.Color = RGB(100, 100, 100)
        .Cells(parameters_start_row + 4, parameters_start_column + 5).NumberFormat = Chr$(34) & "+" & Chr$(34) & "0.00%"
        
        'Weapons Image
        asc_size = 25
        parameters_start_row = parameters_start_row + 4
        For i = 1 To acessories_num
            .Rows(parameters_start_row + i).RowHeight = asc_size
            .Cells(parameters_start_row + i, parameters_start_column + 1).Interior.Color = RGB(8, 46, 84)
            asc_offset_left = .Rows(parameters_start_row + i).ColumnWidth * 5.25/ 2 - asc_size / 2
            asc_offset_top = .Rows(parameters_start_row + i).RowHeight / 2 - asc_size / 2
            Set temp_pic = .Shapes.AddPicture(folder_Path & "\Images\Ascensions\Acessories\" & CStr(i) & ".png", True, True, .Cells(parameters_start_row + i, parameters_start_column + 1).Left _
            + asc_offset_left, .Cells(parameters_start_row + i, parameters_start_column + 1).Top + asc_offset_top, asc_size, asc_size)
            For j = 1 To 4
            .Cells(parameters_start_row + i, parameters_start_column + 1 + j).Value = 0
            .Cells(parameters_start_row + i, parameters_start_column + 1 + j).Interior.Color = RGB(255, 0, 0)
            Next j
        Next i

        'Quality Indicator
        parameters_start_row = 29
        parameters_start_column = 9
        'Title
        .Range(.Cells(parameters_start_row, parameters_start_column + 1), .Cells(parameters_start_row , parameters_start_column + 5)).Merge
        .Cells(parameters_start_row, parameters_start_column + 1).Interior.Color =  RGB(35, 35, 35)
        .Cells(parameters_start_row, parameters_start_column + 1).Value = "Base Quality Chance"
        quality_size = 35
        .Rows(parameters_start_row + 1).RowHeight = 40
        quality_offset_left = .Rows(parameters_start_row + 1).ColumnWidth * 5.25/ 2 - quality_size / 2
        quality_offset_top = .Rows(parameters_start_row + 1).RowHeight / 2 - quality_size / 2
        For i = 1 To 5
        Set temp_pic = .Shapes.AddPicture(folder_Path & "\Images\Quality_Indicators\" & CStr(i) & ".png", True, True, .Cells(parameters_start_row + 1, parameters_start_column + i).Left _
        + quality_offset_left, .Cells(parameters_start_row + 1, parameters_start_column + i).Top + quality_offset_top, quality_size, quality_size)
        .Cells(parameters_start_row + 2, parameters_start_column + i).NumberFormat = "0.00%"
        .Cells(parameters_start_row + 1, parameters_start_column + i).Interior.Color = RGB(55, 55, 55)
        .Cells(parameters_start_row + 2, parameters_start_column + i).Interior.Color = RGB(255, 0, 0)
        Next i
        'Distribution
        .Cells(parameters_start_row + 2, parameters_start_column + 5).Value = 0.002
        .Cells(parameters_start_row + 2, parameters_start_column + 4).Value = 0.004
        .Cells(parameters_start_row + 2, parameters_start_column + 3).Value = 0.012
        .Cells(parameters_start_row + 2, parameters_start_column + 2).Value = 0.06
        .Cells(parameters_start_row + 2, parameters_start_column + 1).Value = 0.922
        'protect the image
        .Protect DrawingObjects:=True, Contents:=False
    End With

End Sub





















