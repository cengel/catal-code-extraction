Option Compare Database
Option Explicit

Private Sub cmdAbrasion_Click()
On Error GoTo err_Ab

    DoCmd.OpenForm "Frm_Admin_Abrasion", acNormal
Exit Sub

err_Ab:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdBaseSection_Click()
On Error GoTo err_Base

    DoCmd.OpenForm "Frm_Admin_BaseSection", acNormal
Exit Sub

err_Base:
    Call General_Error_Trap
    Exit Sub

End Sub

Private Sub cmdClose_Click()
On Error GoTo err_close

DoCmd.OpenForm "Frm_Menu"
DoCmd.Close acForm, Me.Name


Exit Sub

err_close:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdClues_Click()
On Error GoTo err_Cl

    DoCmd.OpenForm "Frm_Admin_CluesForUse", acNormal
Exit Sub

err_Cl:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdColour_Click()
On Error GoTo err_Clr

    DoCmd.OpenForm "Frm_Admin_Colour", acNormal
Exit Sub

err_Clr:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdCompleteness_Click()
On Error GoTo err_Co

    DoCmd.OpenForm "Frm_Admin_Completeness", acNormal
Exit Sub

err_Co:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdCondition_Click()
On Error GoTo err_con

    DoCmd.OpenForm "Frm_Admin_ConditionDetail", acNormal
Exit Sub

err_con:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdElement_Click()
On Error GoTo err_E

    DoCmd.OpenForm "Frm_Admin_Element", acNormal
Exit Sub

err_E:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdForming_Enter()
On Error GoTo err_Frm

    DoCmd.OpenForm "Frm_Admin_Forming", acNormal
Exit Sub

err_Frm:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdHandleTypes_Click()
On Error GoTo err_Ht

    DoCmd.OpenForm "Frm_Admin_HandleTypes", acNormal
Exit Sub

err_Ht:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdInclusions_Click()
On Error GoTo err_I

    DoCmd.OpenForm "Frm_Admin_Inclusions", acNormal
Exit Sub

err_I:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdLetterCode_Click()
On Error GoTo err_Lc

    DoCmd.OpenForm "Frm_Admin_LetterCode", acNormal
Exit Sub

err_Lc:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdPeriods_Click()
On Error GoTo err_P

    DoCmd.OpenForm "Frm_Admin_Period", acNormal
Exit Sub

err_P:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdRimSection_Click()
On Error GoTo err_Rim

    DoCmd.OpenForm "Frm_Admin_RimSection", acNormal
Exit Sub

err_Rim:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdSecondary_Click()
On Error GoTo err_Sec

    DoCmd.OpenForm "Frm_Admin_SecondaryUse", acNormal
Exit Sub

err_Sec:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdSherdSection_Click()
On Error GoTo err_SS

    DoCmd.OpenForm "Frm_Admin_SherdSection", acNormal
Exit Sub

err_SS:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdSTreatment_Click()
On Error GoTo err_ST

    DoCmd.OpenForm "Frm_Admin_SurfaceTreatment", acNormal
Exit Sub

err_ST:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdtechnology_Click()
On Error GoTo err_T

    DoCmd.OpenForm "Frm_Admin_Technology", acNormal
Exit Sub

err_T:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdTemper_Click()

On Error GoTo err_T

    DoCmd.OpenForm "Frm_Admin_Temper", acNormal
Exit Sub

err_T:
    Call General_Error_Trap
    Exit Sub
End Sub


Private Sub cmdWareCode_Click()
On Error GoTo err_Wc

    DoCmd.OpenForm "Frm_Admin_WareGroup", acNormal
Exit Sub

err_Wc:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdTypecodeBases_Click()

On Error GoTo err_cmdTypecodeBases

    DoCmd.OpenForm "Frm_Admin_Typecodes_Bases"
    DoCmd.Maximize
Exit Sub

err_cmdTypecodeBases:
    Call General_Error_Trap
    Exit Sub


End Sub

Private Sub cmdTypecodeDecorations_Click()
On Error GoTo err_cmdTypecodeDecorations

    DoCmd.OpenForm "Frm_Admin_Typecodes_Decorations"
    DoCmd.Maximize
Exit Sub

err_cmdTypecodeDecorations:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdTypecodeHandles_Click()
On Error GoTo err_cmdTypecodeHandles

    DoCmd.OpenForm "Frm_Admin_Typecodes_Handles"
    DoCmd.Maximize
Exit Sub

err_cmdTypecodeHandles:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdTypecodeKnobs_Click()
On Error GoTo err_cmdTypecodeKnobs

    DoCmd.OpenForm "Frm_Admin_Typecodes_Knobs"
    DoCmd.Maximize
Exit Sub

err_cmdTypecodeKnobs:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdTypecodeLids_Click()
On Error GoTo err_cmdTypecodeLids

    DoCmd.OpenForm "Frm_Admin_Typecodes_Lids"
    DoCmd.Maximize
Exit Sub

err_cmdTypecodeLids:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdTypecodeLugs_Click()
On Error GoTo err_cmdTypecodeLugs

    DoCmd.OpenForm "Frm_Admin_Typecodes_Lugs"
    DoCmd.Maximize
Exit Sub

err_cmdTypecodeLugs:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdTypecodeRareForms_Click()
On Error GoTo err_cmdTypecodeRareForms

    DoCmd.OpenForm "Frm_Admin_Typecodes_RareForms"
    DoCmd.Maximize
Exit Sub

err_cmdTypecodeRareForms:
    Call General_Error_Trap
    Exit Sub
End Sub

Private Sub cmdTypecodeRims_Click()

On Error GoTo err_cmdTypecodeRims

    DoCmd.OpenForm "Frm_Admin_Typecodes_Rims"
    DoCmd.Maximize
Exit Sub

err_cmdTypecodeRims:
    Call General_Error_Trap
    Exit Sub


End Sub
