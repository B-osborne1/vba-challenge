# vba-challenge

Overview: This repository involves a tri-year examination of different stocks from 2018-2020. A macro was established to summarise how each stock changed per year and to identify the most significant changes.

Section 1: Contents
    1) Summary_Screenshots: A screenshot for 2018, 2019 and 2020 outlining results
    2) Stock_Analysis_Macro.vba: The macro developed to complete task

Section 2: Sources Used

    Note 1: How to loop through all worksheets
        Source: Excel Destination (content creator) 2019
                YouTube: https://www.youtube.com/watch?v=AlC8a7KyJq0  -- @2:20

                    Dim a As Integer
                    a = Application.Worksheets.Count
                    For J = 1 To a
                    Worksheets(J).Activate

    Note 2: How to autofit all cells
        Source: Puneet 2023 
                Excel Champes: https://excelchamps.com/vba/autofit/

                    ActiveSheet.UsedRange.EntireColumn.AutoFit
                    ActiveSheet.UsedRange.EntireRow.AutoFit

                    
