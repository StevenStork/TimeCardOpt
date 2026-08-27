Attribute VB_Name = "modTemplateReset"
Option Explicit

'==============================================================================
' Template Reset
'
' Clears all project/direct charge codes and Time Entry hour entries while
' leaving Welcome!B3 and Welcome!F3 unchanged.
'
' Assign ResetTemplate to a Welcome-sheet button named ResetButton, or run
' from the Macros dialog. UpdateWelcomeSummary also stacks ResetButton in
' column R with the other action buttons.
'==============================================================================

'------------------------------------------------------------------------------
Public Sub ResetTemplate()
    On Error GoTo Fail

    If MsgBox( _
        "Reset this template?" & vbCrLf & vbCrLf & _
        "This will remove all direct charge codes and projects, and clear " & _
        "all Time Entry hour records." & vbCrLf & vbCrLf & _
        "Welcome!B3 and Welcome!F3 will not be changed.", _
        vbYesNo + vbExclamation + vbDefaultButton2, "Reset Template") <> vbYes Then
        Exit Sub
    End If

    OptimizeExcel True
    ClearAllProjectCodes
    ClearTimeEntryCharges
    OptimizeExcel False

    UpdateWelcomeSummary

    MsgBox "Template reset complete.", vbInformation, "Reset Template"
    Exit Sub

Fail:
    OptimizeExcel False
    MsgBox "ResetTemplate failed: " & Err.Description, vbExclamation, "Reset Template"
End Sub
