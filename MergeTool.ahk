#Requires AutoHotkey v2.0
#SingleInstance Force

;@Ahk2Exe-SetMainIcon MergeTool.ico
;@Ahk2Exe-SetName MergeTool
;@Ahk2Exe-SetDescription MergeTool
;@Ahk2Exe-SetVersion 1.0.0

if !A_IsCompiled && FileExist(A_ScriptDir "\MergeTool.ico")
    TraySetIcon(A_ScriptDir "\MergeTool.ico")

; ========================================
; MERGE TOOL
; ========================================
; A comprehensive merge tool application integrating Word, Excel, and Outlook
; via COM automation for sending emails, individual documents, and printer output.
;
; FEATURES:
; • Modern GUI with dynamic options panel
; • Three merge types: Email, Documents, Print
; • Excel data source integration
; • Field selection from data source
; • Progress tracking and status updates
; • User preference persistence
;
; ========================================

; ========================================
; Constants
; ========================================

; Word MailMerge Destination constants
global WD_SEND_TO_NEW_DOCUMENT := 0
global WD_SEND_TO_PRINTER := 1
global WD_SEND_TO_EMAIL := 2
global WD_SEND_TO_FAX := 3

; Word MailMerge Type constants
global WD_FORM_LETTERS := 0
global WD_MAILING_LABELS := 1
global WD_ENVELOPES := 2
global WD_CATALOG := 3

; Word SaveAs Format constants
global WD_FORMAT_DOCUMENT := 0
global WD_FORMAT_DOCUMENT_DEFAULT := 16
global WD_FORMAT_PDF := 17
global WD_FORMAT_HTML := 8
global WD_FORMAT_FILTERED_HTML := 10  ; Cleaner HTML for email
global WD_FORMAT_RTF := 6

; Excel SaveAs Format constants
global XL_OPENXML_WORKBOOK := 51  ; .xlsx format

; Word ExportAsFixedFormat constants
global WD_EXPORT_FORMAT_PDF := 17            ; Export as PDF
global WD_EXPORT_OPTIMIZE_FOR_PRINT := 0     ; Optimize for printing
global WD_EXPORT_OPTIMIZE_FOR_SCREEN := 1    ; Optimize for screen
global WD_EXPORT_ALL_DOCUMENT := 0           ; Export entire document
global WD_EXPORT_SELECTION := 1              ; Export selection only
global WD_EXPORT_CURRENT_PAGE := 2           ; Export current page only
global WD_EXPORT_FROM_TO := 3                ; Export page range

; Word Statistics constants
global WD_STATISTIC_PAGES := 2              ; For ComputeStatistics - page count

; Word MailMerge Record Navigation constants
global WD_FIRST_RECORD := -4
global WD_LAST_RECORD := -16
global WD_NEXT_RECORD := -5
global WD_PREVIOUS_RECORD := -6
global WD_DEFAULT_FIRST_RECORD := 1
global WD_DEFAULT_LAST_RECORD := -16

; Word Document Close constants
global WD_DO_NOT_SAVE_CHANGES := 0
global WD_SAVE_CHANGES := -1
global WD_PROMPT_TO_SAVE_CHANGES := -2

; Word Alert Level constants
global WD_ALERTS_NONE := 0
global WD_ALERTS_MESSAGE_BOX := -2
global WD_ALERTS_ALL := -1

; Outlook constants
global OL_MAIL_ITEM := 0
global OL_BODY_FORMAT_PLAIN := 1
global OL_BODY_FORMAT_HTML := 2
global OL_BODY_FORMAT_RTF := 3

; ========================================
; Color Scheme
; ========================================
global BG_COLOR := "F5F6FA"          ; Light blue-gray background
global ACCENT_COLOR := "4A90E2"      ; Modern blue
global SUCCESS_COLOR := "52C41A"     ; Green
global ERROR_COLOR := "F5222D"       ; Red
global TEXT_COLOR := "2C3E50"        ; Dark gray text
global SECTION_BG := "FFFFFF"        ; White for sections

; ========================================
; Configuration
; ========================================
global CONFIG_FILE := A_ScriptDir "\MergeTool_config.ini"

; ========================================
; GUI Controls (Global)
; ========================================
global MainGui
global TemplatePathEdit, DataSourcePathEdit
global MergeTypeEmail, MergeTypeDocuments, MergeTypePrint
global StatusText, ProgressBar1, DetailedProgress
global StartMergeBtn, CancelBtn, LoadFieldsBtn

; Dynamic Panel Controls - Email
global EmailPanel
global EmailFieldDropdown, SubjectFieldDropdown
global CCFieldDropdown, CCField2Dropdown, CCField3Dropdown
global BCCFieldDropdown, BCCField2Dropdown, BCCField3Dropdown
global AttachmentField1DD, AttachmentField2DD, AttachmentField3DD
global CommonAttachmentEdit, CommonAttachmentPath := ""
global SendFromAccountDropdown
global DelaySecondsEdit

; Dynamic Panel Controls - Document
global DocumentPanel
global FilenameFieldDropdown, FilenamePrefixEdit
global OutputFolderEdit

; Dynamic Panel Controls - Print
global PrintPanel
global PrinterDropdown
global LastSavedPrinter := ""
global LastSavedAccount := ""
global CollateCheckbox
global CopiesEdit

; State variables
global SelectedTemplatePath := ""
global SelectedDataSourcePath := ""
global CurrentMergeType := "email"

; Each options panel has its own natural height; the window and everything
; below the panel follow the active one instead of always reserving room for
; the tallest (email).
global GuiReady := false
global EMAIL_PANEL_BOTTOM := 470
global DOC_PANEL_BOTTOM   := 365
global PRINT_PANEL_BOTTOM := 330
global FieldNames := []
global IsProcessing := false

; ========================================
; Initialize
; ========================================
CreateGUI()
LoadPreferences()
EnumeratePrinters()
EnumerateOutlookAccounts()

; ========================================
; GUI Creation
; ========================================
CreateGUI() {
    global MainGui, TemplatePathEdit, DataSourcePathEdit
    global MergeTypeEmail, MergeTypeDocuments, MergeTypePrint
    global StatusText, ProgressBar1, DetailedProgress
    global StartMergeBtn, CancelBtn, LoadFieldsBtn
    global EmailPanel, DocumentPanel, PrintPanel
    global GuiReady

    MainGui := Gui("-MaximizeBox", "MergeTool")
    MainGui.SetFont("s10", "Segoe UI")
    MainGui.BackColor := BG_COLOR

    ; ===== FILE SELECTION SECTION (y: 20-140) =====
    MainGui.SetFont("s10", "Segoe UI")
    MainGui.Add("GroupBox", "x15 y20 w770 h120", " File Selection ")

    ; Template row
    MainGui.Add("Text", "x30 y45 w100 h28 +0x200", "Template:")
    TemplatePathEdit := MainGui.Add("Edit", "x135 y45 w520 h28 ReadOnly", "")
    browseTemplateBtn := MainGui.Add("Button", "x665 y45 w105 h28", "Browse...")
    browseTemplateBtn.OnEvent("Click", BrowseTemplate)

    ; Data source row
    MainGui.Add("Text", "x30 y80 w100 h28 +0x200", "Data Source:")
    DataSourcePathEdit := MainGui.Add("Edit", "x135 y80 w520 h28 ReadOnly", "")
    browseDataBtn := MainGui.Add("Button", "x665 y80 w105 h28", "Browse...")
    browseDataBtn.OnEvent("Click", BrowseDataSource)

    ; Load Fields button
    LoadFieldsBtn := MainGui.Add("Button", "x135 y112 w120 h26", "Load Fields")
    LoadFieldsBtn.OnEvent("Click", LoadFields)

    ; Generate Contact Sheet button
    GenerateSheetBtn := MainGui.Add("Button", "x270 y112 w220 h26", "Generate Contact Sheet...")
    GenerateSheetBtn.OnEvent("Click", GenerateContactSheet)
    GenerateSheetBtn.ToolTip := "Create a new Excel sheet with starter columns (Name, Email, Subject, CC, BCC, Attachment) for your contact list"

    ; ===== MERGE TYPE SELECTION (y: 200-260) =====
    MainGui.Add("GroupBox", "x15 y150 w770 h65", " Merge Type ")

    MergeTypeEmail := MainGui.Add("Radio", "x30 y175 w180 h28 Checked", "Email Merge")
    MergeTypeEmail.OnEvent("Click", (*) => SwitchMergeType("email"))

    MergeTypeDocuments := MainGui.Add("Radio", "x230 y175 w180 h28", "Document Merge")
    MergeTypeDocuments.OnEvent("Click", (*) => SwitchMergeType("documents"))

    MergeTypePrint := MainGui.Add("Radio", "x430 y175 w180 h28", "Printer Merge")
    MergeTypePrint.OnEvent("Click", (*) => SwitchMergeType("print"))

    ; ===== DYNAMIC OPTIONS PANEL (y: 275-540) =====
    ; Create all three panels, only show the active one

    CreateEmailPanel()
    CreateDocumentPanel()
    CreatePrintPanel()

    ; Show email panel by default
    ShowPanel("email")

    ; ===== ACTION BUTTONS (y: 560-605) =====
    MainGui.SetFont("s11 Bold", "Segoe UI")
    StartMergeBtn := MainGui.Add("Button", "x15 y482 w600 h32 Default", "Start Merge")
    StartMergeBtn.OnEvent("Click", StartMerge)

    CancelBtn := MainGui.Add("Button", "x625 y482 w160 h32", "Cancel")
    CancelBtn.OnEvent("Click", CancelMerge)
    CancelBtn.Enabled := false

    ; ===== PROGRESS & STATUS SECTION (y: 526-586) =====
    MainGui.SetFont("s10", "Segoe UI")
    ProgressBar1 := MainGui.Add("Progress", "x15 y526 w770 h16 BackgroundE8E8E8 c" . ACCENT_COLOR, 0)

    StatusText := MainGui.Add("Text", "x15 y548 w770 h20", "Ready - Select template and data source to begin")
    StatusText.SetFont("s9 bold")

    DetailedProgress := MainGui.Add("Text", "x15 y568 w770 h18 c888888", "")

    ; ===== TOOLTIPS =====
    TemplatePathEdit.ToolTip := "Word document (.doc or .docx) containing merge fields"
    DataSourcePathEdit.ToolTip := "Excel workbook (.xls or .xlsx) containing data records"
    LoadFieldsBtn.ToolTip := "Load field names from the data source"
    StartMergeBtn.ToolTip := "Start the mail merge process"
    CancelBtn.ToolTip := "Cancel the current merge operation"
    MergeTypeEmail.ToolTip := "Send emails to Outlook Outbox with attachments or HTML body"
    MergeTypeDocuments.ToolTip := "Create individual PDF files for each record"
    MergeTypePrint.ToolTip := "Send merged documents directly to a printer"

    ; ===== EVENT HANDLERS =====
    MainGui.OnEvent("Close", (*) => ExitApp())

    ; Show GUI
    MainGui.Show("w800 h598")
    GuiReady := true
}

; ========================================
; Email Panel Creation
; ========================================
CreateEmailPanel() {
    global MainGui, EmailPanel
    global EmailFieldDropdown, SubjectFieldDropdown
    global CCFieldDropdown, CCField2Dropdown, CCField3Dropdown
    global BCCFieldDropdown, BCCField2Dropdown, BCCField3Dropdown
    global AttachmentField1DD, AttachmentField2DD, AttachmentField3DD
    global CommonAttachmentEdit
    global SendFromAccountDropdown
    global DelaySecondsEdit

    ; Email panel container (using GroupBox) - height adjusted for CC/BCC fields
    EmailPanel := MainGui.Add("GroupBox", "x15 y225 w770 h245 vEmailPanel", " Email Options ")

    ; Send From account selection
    MainGui.Add("Text", "x30 y250 w120 h28 +0x200 vSendFromLabel", "Send From:")
    SendFromAccountDropdown := MainGui.Add("DropDownList", "x155 y250 w400 h200 vSendFromAccountDD", ["(Default Account)"])
    SendFromAccountDropdown.ToolTip := "Select which Outlook account to send emails from"
    SendFromAccountDropdown.OnEvent("Change", SaveAccountPreference)
    refreshAccountsBtn := MainGui.Add("Button", "x565 y250 w70 h28 vRefreshAccountsBtn", "Refresh")
    refreshAccountsBtn.OnEvent("Click", (*) => EnumerateOutlookAccounts())
    refreshAccountsBtn.ToolTip := "Refresh the list of available Outlook accounts"

    ; Delay between sends (helps avoid spam/rate-limit blocks from mail servers)
    MainGui.Add("Text", "x640 y250 w55 h28 +0x200 vDelayLabel", "Delay:")
    DelaySecondsEdit := MainGui.Add("Edit", "x695 y250 w40 h28 Number vDelaySecondsED", "15")
    MainGui.Add("UpDown", "vDelaySecondsUD Range0-600", 15)
    DelaySecondsEdit.ToolTip := "Seconds to wait between each email send - helps avoid triggering spam/rate-limit blocks from your mail server"
    DelaySecondsEdit.OnEvent("Change", SaveDelayPreference)

    ; Email field
    MainGui.Add("Text", "x30 y280 w120 h28 +0x200 vEmailLabel", "Email Field:")
    EmailFieldDropdown := MainGui.Add("DropDownList", "x155 y280 w250 h200 vEmailFieldDD", ["(Load fields first)"])
    EmailFieldDropdown.ToolTip := "Select the field containing email addresses"

    ; Subject options
    MainGui.Add("Text", "x30 y310 w120 h28 +0x200 vSubjectLabel", "Subject:")
    SubjectFieldDropdown := MainGui.Add("DropDownList", "x155 y310 w250 h200 vSubjectFieldDD", ["(Load fields first)"])
    SubjectFieldDropdown.ToolTip := "Select the field containing email subjects"

    ; CC fields (3 dropdowns for data source columns)
    MainGui.Add("Text", "x30 y340 w120 h28 +0x200 vCCLabel", "CC Fields:")
    CCFieldDropdown := MainGui.Add("DropDownList", "x155 y340 w170 h200 vCCFieldDD", ["(None)", "(Load fields first)"])
    CCFieldDropdown.ToolTip := "Optional: Select field 1 containing CC addresses"
    CCField2Dropdown := MainGui.Add("DropDownList", "x335 y340 w170 h200 vCCField2DD", ["(None)"])
    CCField2Dropdown.ToolTip := "Optional: Select field 2 containing CC addresses"
    CCField3Dropdown := MainGui.Add("DropDownList", "x515 y340 w170 h200 vCCField3DD", ["(None)"])
    CCField3Dropdown.ToolTip := "Optional: Select field 3 containing CC addresses"

    ; BCC fields (3 dropdowns for data source columns)
    MainGui.Add("Text", "x30 y370 w120 h28 +0x200 vBCCLabel", "BCC Fields:")
    BCCFieldDropdown := MainGui.Add("DropDownList", "x155 y370 w170 h200 vBCCFieldDD", ["(None)", "(Load fields first)"])
    BCCFieldDropdown.ToolTip := "Optional: Select field 1 containing BCC addresses"
    BCCField2Dropdown := MainGui.Add("DropDownList", "x335 y370 w170 h200 vBCCField2DD", ["(None)"])
    BCCField2Dropdown.ToolTip := "Optional: Select field 2 containing BCC addresses"
    BCCField3Dropdown := MainGui.Add("DropDownList", "x515 y370 w170 h200 vBCCField3DD", ["(None)"])
    BCCField3Dropdown.ToolTip := "Optional: Select field 3 containing BCC addresses"

    ; Per-person attachment fields (from data source columns)
    MainGui.Add("Text", "x30 y400 w120 h28 +0x200 vAttachFieldLabel", "Attach Fields:")
    AttachmentField1DD := MainGui.Add("DropDownList", "x155 y400 w170 h200 vAttachField1DD", ["(None)"])
    AttachmentField1DD.ToolTip := "Select column containing file path for attachment 1"
    AttachmentField2DD := MainGui.Add("DropDownList", "x335 y400 w170 h200 vAttachField2DD", ["(None)"])
    AttachmentField2DD.ToolTip := "Select column containing file path for attachment 2"
    AttachmentField3DD := MainGui.Add("DropDownList", "x515 y400 w170 h200 vAttachField3DD", ["(None)"])
    AttachmentField3DD.ToolTip := "Select column containing file path for attachment 3"

    ; Common attachment for all emails
    MainGui.Add("Text", "x30 y430 w120 h28 +0x200 vCommonAttachLabel", "Common Attach:")
    CommonAttachmentEdit := MainGui.Add("Edit", "x155 y430 w450 h28 ReadOnly vCommonAttachED", "")
    CommonAttachmentEdit.ToolTip := "File attached to ALL emails"
    browseCommonBtn := MainGui.Add("Button", "x615 y430 w70 h28 vBrowseCommonBtn", "Browse")
    browseCommonBtn.OnEvent("Click", BrowseCommonAttachment)
    browseCommonBtn.ToolTip := "Select a file to attach to all emails"

    ; Info text
}

; ========================================
; Document Panel Creation
; ========================================
CreateDocumentPanel() {
    global MainGui, DocumentPanel
    global FilenameFieldDropdown, FilenamePrefixEdit
    global OutputFolderEdit

    ; Document panel container
    DocumentPanel := MainGui.Add("GroupBox", "x15 y225 w770 h140 vDocumentPanel Hidden", " Document Options ")

    ; Output format
    MainGui.Add("Text", "x30 y250 w120 h28 +0x200 vDocFormatLabel Hidden", "Output Format:")
    MainGui.Add("Text", "x155 y250 w120 h28 +0x200 vDocFormatPDFText Hidden", "PDF (.pdf)")

    ; Filename field
    MainGui.Add("Text", "x30 y285 w120 h28 +0x200 vFilenameLabel Hidden", "Filename Field:")
    FilenameFieldDropdown := MainGui.Add("DropDownList", "x155 y285 w250 h200 vFilenameFieldDD Hidden", ["(Load fields first)"])
    FilenameFieldDropdown.ToolTip := "Select the field to use for naming output files"

    ; Filename prefix
    MainGui.Add("Text", "x420 y285 w80 h28 +0x200 vPrefixLabel Hidden", "Prefix:")
    FilenamePrefixEdit := MainGui.Add("Edit", "x505 y285 w250 h28 vPrefixEdit Hidden", "")
    FilenamePrefixEdit.ToolTip := "Optional prefix to add before each filename"

    ; Output folder
    MainGui.Add("Text", "x30 y325 w120 h28 +0x200 vOutputLabel Hidden", "Output Folder:")
    OutputFolderEdit := MainGui.Add("Edit", "x155 y325 w520 h28 ReadOnly vOutputFolderED Hidden", "")
    OutputFolderEdit.ToolTip := "Folder where merged documents will be saved"
    browseOutputBtn := MainGui.Add("Button", "x685 y325 w70 h28 vBrowseOutputBtn Hidden", "Browse")
    browseOutputBtn.OnEvent("Click", BrowseOutputFolder)
    browseOutputBtn.ToolTip := "Select the output folder for merged documents"

    ; Info text
}

; ========================================
; Print Panel Creation
; ========================================
CreatePrintPanel() {
    global MainGui, PrintPanel
    global PrinterDropdown, CollateCheckbox, CopiesEdit

    ; Print panel container
    PrintPanel := MainGui.Add("GroupBox", "x15 y225 w770 h105 vPrintPanel Hidden", " Print Options ")

    ; Printer selection
    MainGui.Add("Text", "x30 y250 w120 h28 +0x200 vPrinterLabel Hidden", "Printer:")
    PrinterDropdown := MainGui.Add("DropDownList", "x155 y250 w400 h200 vPrinterDD Hidden", ["(Loading printers...)"])
    PrinterDropdown.ToolTip := "Select the printer to use for printing"
    PrinterDropdown.OnEvent("Change", SavePrinterPreference)

    ; Copies
    MainGui.Add("Text", "x580 y250 w60 h28 +0x200 vCopiesLabel Hidden", "Copies:")
    CopiesEdit := MainGui.Add("Edit", "x645 y250 w60 h28 Number vCopiesED Hidden", "1")
    CopiesEdit.ToolTip := "Number of copies to print for each record"
    MainGui.Add("UpDown", "vCopiesUD Hidden Range1-99", 1)

    ; Print options
    CollateCheckbox := MainGui.Add("Checkbox", "x155 y290 w150 h28 Checked vCollateCB Hidden", "Collate copies")
    CollateCheckbox.ToolTip := "Print complete sets before starting the next copy"

}

; ========================================
; Panel Switching
; ========================================
SwitchMergeType(mergeType) {
    global CurrentMergeType

    CurrentMergeType := mergeType
    ShowPanel(mergeType)
    SavePreference("LastMergeType", mergeType)
    UpdateStatus("Merge type changed to: " . mergeType, "normal")
}

ShowPanel(panelType) {
    ; Hide all panels and their controls
    HideEmailPanel()
    HideDocumentPanel()
    HidePrintPanel()

    ; Show the selected panel
    switch panelType {
        case "email":
            ShowEmailPanel()
        case "documents":
            ShowDocumentPanel()
        case "print":
            ShowPrintPanel()
    }

    LayoutForPanel(panelType)
}

/**
 * Moves the action buttons and progress area to sit under the active panel and
 * resizes the window to match. Called on every panel switch.
 * @param panelType - "email", "documents" or "print"
 */
LayoutForPanel(panelType) {
    global MainGui, GuiReady
    global StartMergeBtn, CancelBtn, ProgressBar1, StatusText, DetailedProgress
    global EMAIL_PANEL_BOTTOM, DOC_PANEL_BOTTOM, PRINT_PANEL_BOTTOM

    ; ShowPanel runs once during construction, before the action buttons and
    ; progress controls exist. Nothing to lay out until the window is up.
    if !GuiReady
        return

    switch panelType {
        case "documents": panelBottom := DOC_PANEL_BOTTOM
        case "print":     panelBottom := PRINT_PANEL_BOTTOM
        default:          panelBottom := EMAIL_PANEL_BOTTOM
    }

    StartMergeBtn.Move(, panelBottom + 12)
    CancelBtn.Move(, panelBottom + 12)
    ProgressBar1.Move(, panelBottom + 56)
    StatusText.Move(, panelBottom + 78)
    DetailedProgress.Move(, panelBottom + 98)

    MainGui.Show("w800 h" . (panelBottom + 128))
}

ShowEmailPanel() {
    global MainGui

    ; Show email panel and all its controls
    MainGui["EmailPanel"].Visible := true
    MainGui["SendFromLabel"].Visible := true
    MainGui["SendFromAccountDD"].Visible := true
    MainGui["RefreshAccountsBtn"].Visible := true
    MainGui["DelayLabel"].Visible := true
    MainGui["DelaySecondsED"].Visible := true
    MainGui["DelaySecondsUD"].Visible := true
    MainGui["EmailLabel"].Visible := true
    MainGui["EmailFieldDD"].Visible := true
    MainGui["SubjectLabel"].Visible := true
    MainGui["SubjectFieldDD"].Visible := true
    MainGui["CCLabel"].Visible := true
    MainGui["CCFieldDD"].Visible := true
    MainGui["CCField2DD"].Visible := true
    MainGui["CCField3DD"].Visible := true
    MainGui["BCCLabel"].Visible := true
    MainGui["BCCFieldDD"].Visible := true
    MainGui["BCCField2DD"].Visible := true
    MainGui["BCCField3DD"].Visible := true
    MainGui["AttachFieldLabel"].Visible := true
    MainGui["AttachField1DD"].Visible := true
    MainGui["AttachField2DD"].Visible := true
    MainGui["AttachField3DD"].Visible := true
    MainGui["CommonAttachLabel"].Visible := true
    MainGui["CommonAttachED"].Visible := true
    MainGui["BrowseCommonBtn"].Visible := true
}

HideEmailPanel() {
    global MainGui

    MainGui["EmailPanel"].Visible := false
    MainGui["SendFromLabel"].Visible := false
    MainGui["SendFromAccountDD"].Visible := false
    MainGui["RefreshAccountsBtn"].Visible := false
    MainGui["DelayLabel"].Visible := false
    MainGui["DelaySecondsED"].Visible := false
    MainGui["DelaySecondsUD"].Visible := false
    MainGui["EmailLabel"].Visible := false
    MainGui["EmailFieldDD"].Visible := false
    MainGui["SubjectLabel"].Visible := false
    MainGui["SubjectFieldDD"].Visible := false
    MainGui["CCLabel"].Visible := false
    MainGui["CCFieldDD"].Visible := false
    MainGui["CCField2DD"].Visible := false
    MainGui["CCField3DD"].Visible := false
    MainGui["BCCLabel"].Visible := false
    MainGui["BCCFieldDD"].Visible := false
    MainGui["BCCField2DD"].Visible := false
    MainGui["BCCField3DD"].Visible := false
    MainGui["AttachFieldLabel"].Visible := false
    MainGui["AttachField1DD"].Visible := false
    MainGui["AttachField2DD"].Visible := false
    MainGui["AttachField3DD"].Visible := false
    MainGui["CommonAttachLabel"].Visible := false
    MainGui["CommonAttachED"].Visible := false
    MainGui["BrowseCommonBtn"].Visible := false
}

ShowDocumentPanel() {
    global MainGui

    MainGui["DocumentPanel"].Visible := true
    MainGui["DocFormatLabel"].Visible := true
    MainGui["DocFormatPDFText"].Visible := true
    MainGui["FilenameLabel"].Visible := true
    MainGui["FilenameFieldDD"].Visible := true
    MainGui["PrefixLabel"].Visible := true
    MainGui["PrefixEdit"].Visible := true
    MainGui["OutputLabel"].Visible := true
    MainGui["OutputFolderED"].Visible := true
    MainGui["BrowseOutputBtn"].Visible := true
}

HideDocumentPanel() {
    global MainGui

    MainGui["DocumentPanel"].Visible := false
    MainGui["DocFormatLabel"].Visible := false
    MainGui["DocFormatPDFText"].Visible := false
    MainGui["FilenameLabel"].Visible := false
    MainGui["FilenameFieldDD"].Visible := false
    MainGui["PrefixLabel"].Visible := false
    MainGui["PrefixEdit"].Visible := false
    MainGui["OutputLabel"].Visible := false
    MainGui["OutputFolderED"].Visible := false
    MainGui["BrowseOutputBtn"].Visible := false
}

ShowPrintPanel() {
    global MainGui

    MainGui["PrintPanel"].Visible := true
    MainGui["PrinterLabel"].Visible := true
    MainGui["PrinterDD"].Visible := true
    MainGui["CopiesLabel"].Visible := true
    MainGui["CopiesED"].Visible := true
    MainGui["CopiesUD"].Visible := true
    MainGui["CollateCB"].Visible := true
}

HidePrintPanel() {
    global MainGui

    MainGui["PrintPanel"].Visible := false
    MainGui["PrinterLabel"].Visible := false
    MainGui["PrinterDD"].Visible := false
    MainGui["CopiesLabel"].Visible := false
    MainGui["CopiesED"].Visible := false
    MainGui["CopiesUD"].Visible := false
    MainGui["CollateCB"].Visible := false
}

; ========================================
; File Browser Functions
; ========================================
BrowseTemplate(*) {
    global SelectedTemplatePath, TemplatePathEdit

    selectedFile := FileSelect(3, "", "Select Word Template", "Word Documents (*.docx; *.doc)")
    if selectedFile {
        SelectedTemplatePath := selectedFile
        TemplatePathEdit.Value := selectedFile
        SavePreference("LastTemplate", selectedFile)
        UpdateStatus("Template selected: " . selectedFile, "success")
    }
}

BrowseDataSource(*) {
    global SelectedDataSourcePath, DataSourcePathEdit

    selectedFile := FileSelect(3, "", "Select Excel Data Source", "Excel Workbooks (*.xlsx; *.xls)")
    if selectedFile {
        SelectedDataSourcePath := selectedFile
        DataSourcePathEdit.Value := selectedFile
        SavePreference("LastDataSource", selectedFile)
        UpdateStatus("Data source selected: " . selectedFile, "success")
    }
}

BrowseOutputFolder(*) {
    global OutputFolderEdit

    selectedFolder := DirSelect(, 3, "Select Output Folder")
    if selectedFolder {
        OutputFolderEdit.Value := selectedFolder
        SavePreference("LastOutputFolder", selectedFolder)
        UpdateStatus("Output folder selected: " . selectedFolder, "success")
    }
}

BrowseCommonAttachment(*) {
    global CommonAttachmentEdit, CommonAttachmentPath

    selectedFile := FileSelect(3, "", "Select Common Attachment", "All Files (*.*)")
    if selectedFile {
        CommonAttachmentPath := selectedFile
        CommonAttachmentEdit.Value := selectedFile
        UpdateStatus("Common attachment selected: " . selectedFile, "success")
    }
}

; ========================================
; Generate Contact Sheet
; ========================================
GenerateContactSheet(*) {
    global SelectedDataSourcePath, DataSourcePathEdit, XL_OPENXML_WORKBOOK

    savePath := FileSelect("S", A_MyDocuments . "\Contacts.xlsx", "Save Contact Sheet As", "Excel Workbook (*.xlsx)")
    if !savePath
        return

    ; FileSelect doesn't always append the extension
    if !RegExMatch(savePath, "i)\.xlsx$")
        savePath .= ".xlsx"

    UpdateStatus("Creating contact sheet...", "processing", 20)

    excelApp := ""
    workbook := ""

    try {
        excelApp := ComObject("Excel.Application")
        excelApp.Visible := false
        excelApp.DisplayAlerts := false

        workbook := excelApp.Workbooks.Add()
        sheet := workbook.Sheets.Item(1)

        headers := ["Name", "Email", "Subject", "CC", "BCC", "Attachment"]
        for i, header in headers
            sheet.Cells(1, i).Value := header

        workbook.SaveAs(savePath, XL_OPENXML_WORKBOOK)
        workbook.Close(false)
        workbook := ""
        excelApp.Quit()
        excelApp := ""

        ; Open the new sheet so the user can start typing contacts right away
        Run(savePath)

        ; Auto-select it as the Data Source
        SelectedDataSourcePath := savePath
        DataSourcePathEdit.Value := savePath
        SavePreference("LastDataSource", savePath)

        UpdateStatus("Contact sheet created: " . savePath, "success", 100)

    } catch as err {
        UpdateStatus("Error: " . err.Message, "error", 0)
        MsgBox("Error creating contact sheet:`n`n" . err.Message, "Error", "IconX")

    } finally {
        try {
            if workbook != ""
                workbook.Close(false)
        }
        try {
            if excelApp != ""
                excelApp.Quit()
        }
    }
}

; ========================================
; Load Fields from Data Source
; ========================================
LoadFields(*) {
    global SelectedTemplatePath, SelectedDataSourcePath, FieldNames

    ; Validate inputs
    if !SelectedTemplatePath || SelectedTemplatePath = "" {
        MsgBox("Please select a Word template first.", "No Template", "Icon!")
        UpdateStatus("Error: No template selected", "error")
        return
    }

    if !SelectedDataSourcePath || SelectedDataSourcePath = "" {
        MsgBox("Please select an Excel data source first.", "No Data Source", "Icon!")
        UpdateStatus("Error: No data source selected", "error")
        return
    }

    if !FileExist(SelectedTemplatePath) {
        MsgBox("Template file not found: " . SelectedTemplatePath, "File Not Found", "Icon!")
        UpdateStatus("Error: Template file not found", "error")
        return
    }

    if !FileExist(SelectedDataSourcePath) {
        MsgBox("Data source file not found: " . SelectedDataSourcePath, "File Not Found", "Icon!")
        UpdateStatus("Error: Data source file not found", "error")
        return
    }

    UpdateStatus("Loading fields from data source...", "processing", 10)
    SetControlsEnabled(false)

    wordApp := ""
    doc := ""

    try {
        ; Initialize Word COM application
        UpdateStatus("Initializing Word...", "processing", 20)
        wordApp := ComObject("Word.Application")
        wordApp.Visible := false
        wordApp.DisplayAlerts := 0  ; Suppress alerts

        ; Open the template document
        UpdateStatus("Opening template...", "processing", 35)
        doc := wordApp.Documents.Open(SelectedTemplatePath, false, true)  ; ReadOnly = true

        ; Connect to Excel data source using the modular function
        UpdateStatus("Connecting to data source...", "processing", 50)
        OpenDataSource(wordApp, doc, SelectedDataSourcePath)

        ; Enumerate fields from the connected data source
        UpdateStatus("Reading field names...", "processing", 70)
        FieldNames := EnumerateFields(doc)

        ; Get record count for display
        recordCount := GetRecordCount(doc)

        ; Close document and Word
        UpdateStatus("Cleaning up...", "processing", 90)
        doc.Close(0)
        doc := ""
        wordApp.Quit()
        wordApp := ""

        ; Populate all field dropdowns
        PopulateFieldDropdowns()

        ; Success message with field and record count
        UpdateStatus("Loaded " . FieldNames.Length . " fields, " . recordCount . " records", "success", 100)
        UpdateDetailedProgress("Fields: " . JoinArray(FieldNames, ", "))

    } catch as err {
        UpdateStatus("Error: " . err.Message, "error", 0)
        MsgBox("Error loading fields:`n`n" . err.Message . "`n`nMake sure:`n• Microsoft Word is installed`n• The Excel file is not open in another program`n• The files are valid", "Error", "IconX")

    } finally {
        ; Always cleanup COM objects
        try {
            if doc != ""
                doc.Close(0)
        }
        try {
            if wordApp != ""
                wordApp.Quit()
        }
        SetControlsEnabled(true)
    }
}

/**
 * Joins an array of strings with a delimiter
 * @param arr - Array to join
 * @param delimiter - String to place between elements
 * @returns Joined string
 */
JoinArray(arr, delimiter := ", ") {
    result := ""
    for i, item in arr {
        if i > 1
            result .= delimiter
        result .= item
    }
    return result
}

; ========================================
; Data Source Functions (Phase 2)
; ========================================

/**
 * Reads the name of the first worksheet in an Excel workbook
 * Opens the file read-only in a separate hidden Excel instance so a copy the
 * user already has open is left untouched.
 * @param dataSourcePath - Full path to the Excel file
 * @returns Sheet name, or "" if Excel is unavailable or the file can't be read
 */
GetFirstSheetName(dataSourcePath) {
    excelApp := ""
    workbook := ""
    sheetName := ""

    try {
        excelApp := ComObject("Excel.Application")
        excelApp.Visible := false
        excelApp.DisplayAlerts := false

        ; UpdateLinks = 0 (never), ReadOnly = true
        workbook := excelApp.Workbooks.Open(dataSourcePath, 0, true)
        sheetName := workbook.Sheets.Item(1).Name

    } catch {
        ; Excel missing or file unreadable - caller falls back to "Sheet1"
        sheetName := ""

    } finally {
        try {
            if workbook != ""
                workbook.Close(false)
        }
        try {
            if excelApp != ""
                excelApp.Quit()
        }
    }

    return sheetName
}

/**
 * Opens an Excel data source and connects it to a Word document's MailMerge
 * @param wordApp - Word Application COM object
 * @param doc - Word Document COM object
 * @param dataSourcePath - Full path to the Excel file
 * @param sheetName - Optional sheet name; when omitted the workbook's first
 *                    sheet is used, so renamed and non-English default tabs
 *                    ("Hoja1", "Tabelle1") work without configuration
 * @returns Boolean indicating success
 */
OpenDataSource(wordApp, doc, dataSourcePath, sheetName := "") {
    if !FileExist(dataSourcePath) {
        throw Error("Data source file not found: " . dataSourcePath)
    }

    ; Resolve the sheet to read from when the caller didn't name one
    if sheetName = "" {
        sheetName := GetFirstSheetName(dataSourcePath)
        if sheetName = ""
            sheetName := "Sheet1"  ; Excel unavailable - assume the usual default
    }

    ; Detect file extension and determine appropriate Extended Properties
    SplitPath(dataSourcePath, , , &fileExt)
    fileExt := StrLower(fileExt)

    ; Determine Extended Properties based on file type
    switch fileExt {
        case "xlsx":
            extProps := "Excel 12.0 Xml;HDR=YES"
        case "xlsm":
            extProps := "Excel 12.0 Macro;HDR=YES"
        case "xlsb":
            extProps := "Excel 12.0;HDR=YES"
        case "xls":
            extProps := "Excel 8.0;HDR=YES"
        default:
            extProps := "Excel 12.0 Xml;HDR=YES"  ; Default to xlsx format
    }

    ; SQL query to select all data from the specified sheet
    sqlStatement := "SELECT * FROM [" . sheetName . "$]"

    ; Method 1: Try simple approach - let Word handle the connection
    ; Parameters: Name, Format, ConfirmConversions, ReadOnly, LinkToSource,
    ;             AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert,
    ;             WritePasswordDocument, WritePasswordTemplate, Connection, SQLStatement
    try {
        doc.MailMerge.OpenDataSource(
            dataSourcePath,     ; Name
            0,                  ; Format (0 = auto-detect)
            false,              ; ConfirmConversions
            true,               ; ReadOnly
            false,              ; LinkToSource
            false,              ; AddToRecentFiles
            "",                 ; PasswordDocument
            "",                 ; PasswordTemplate
            false,              ; Revert
            "",                 ; WritePasswordDocument
            "",                 ; WritePasswordTemplate
            "",                 ; Connection (empty = let Word figure it out)
            sqlStatement        ; SQLStatement
        )
        return true
    } catch as err1 {
        ; Method 2: Try with explicit ACE.OLEDB.16.0 connection string (newer provider)
        try {
            connectionString := "Provider=Microsoft.ACE.OLEDB.16.0;"
                . "Data Source=" . dataSourcePath . ";"
                . "Extended Properties='" . extProps . "'"

            doc.MailMerge.OpenDataSource(
                dataSourcePath,
                0, false, true, false, false, "", "", false, "", "",
                connectionString,
                sqlStatement
            )
            return true
        } catch as err2 {
            ; Method 3: Try with ACE.OLEDB.12.0 (older but more common provider)
            try {
                connectionString := "Provider=Microsoft.ACE.OLEDB.12.0;"
                    . "Data Source=" . dataSourcePath . ";"
                    . "Extended Properties='" . extProps . "'"

                doc.MailMerge.OpenDataSource(
                    dataSourcePath,
                    0, false, true, false, false, "", "", false, "", "",
                    connectionString,
                    sqlStatement
                )
                return true
            } catch as err3 {
                ; Method 4: Try older Jet provider for .xls files
                if (fileExt = "xls") {
                    try {
                        connectionString := "Provider=Microsoft.Jet.OLEDB.4.0;"
                            . "Data Source=" . dataSourcePath . ";"
                            . "Extended Properties='Excel 8.0;HDR=YES'"

                        doc.MailMerge.OpenDataSource(
                            dataSourcePath,
                            0, false, true, false, false, "", "", false, "", "",
                            connectionString,
                            sqlStatement
                        )
                        return true
                    } catch as err4 {
                        throw Error("Failed to open data source.`n`nFile type: ." . fileExt . "`nSheet: " . sheetName . "`n`nTried 4 methods:`n1. " . err1.Message . "`n2. " . err2.Message . "`n3. " . err3.Message . "`n4. " . err4.Message)
                    }
                } else {
                    throw Error("Failed to open data source.`n`nFile type: ." . fileExt . "`nSheet: " . sheetName . "`n`nTried 3 methods:`n1. " . err1.Message . "`n2. " . err2.Message . "`n3. " . err3.Message)
                }
            }
        }
    }
}

/**
 * Enumerates field names from the connected data source
 * @param doc - Word Document COM object with connected data source
 * @returns Array of field name strings
 */
EnumerateFields(doc) {
    fields := []

    if !doc.MailMerge.DataSource.DataFields {
        throw Error("No data source connected to document")
    }

    fieldCount := doc.MailMerge.DataSource.FieldNames.Count

    if fieldCount = 0 {
        throw Error("No fields found in data source")
    }

    Loop fieldCount {
        fieldName := doc.MailMerge.DataSource.FieldNames.Item(A_Index).Name
        fields.Push(fieldName)
    }

    return fields
}

/**
 * Gets the total number of records in the connected data source
 * @param doc - Word Document COM object with connected data source
 * @returns Integer record count
 */
GetRecordCount(doc) {
    try {
        ; Method 1: Try using RecordCount property directly
        try {
            recordCount := doc.MailMerge.DataSource.RecordCount
            if recordCount > 0
                return recordCount
        }

        ; Method 2: Count by iterating through records
        ; Start at first record
        doc.MailMerge.DataSource.ActiveRecord := 1
        count := 1

        ; Keep moving to next record until we can't anymore
        Loop {
            try {
                doc.MailMerge.DataSource.ActiveRecord := count + 1
                ; If we successfully moved, increment count
                if doc.MailMerge.DataSource.ActiveRecord = count + 1
                    count++
                else
                    break
            } catch {
                break
            }

            ; Safety limit
            if count > 10000
                break
        }

        ; Reset to first record
        doc.MailMerge.DataSource.ActiveRecord := 1

        return count
    } catch as err {
        throw Error("Failed to get record count: " . err.Message)
    }
}

/**
 * Gets the value of a specific field from the current active record
 * @param doc - Word Document COM object with connected data source
 * @param fieldName - Name of the field to retrieve
 * @returns String value of the field
 */
GetFieldValue(doc, fieldName) {
    try {
        return doc.MailMerge.DataSource.DataFields.Item(fieldName).Value
    } catch as err {
        throw Error("Failed to get field value '" . fieldName . "': " . err.Message)
    }
}

/**
 * Sets the active record to a specific index
 * @param doc - Word Document COM object with connected data source
 * @param recordIndex - 1-based index of the record to select
 */
SetActiveRecord(doc, recordIndex) {
    try {
        doc.MailMerge.DataSource.ActiveRecord := recordIndex
    } catch as err {
        throw Error("Failed to set active record to " . recordIndex . ": " . err.Message)
    }
}

/**
 * Initializes Word application and opens template with data source connected
 * Returns an object with wordApp, doc, and recordCount properties
 * Caller is responsible for cleanup via CleanupWordSession()
 * @param templatePath - Full path to the Word template
 * @param dataSourcePath - Full path to the Excel data source
 * @param showWord - Whether to make Word visible (default: false)
 * @returns Object {wordApp, doc, recordCount}
 */
InitializeWordSession(templatePath, dataSourcePath, showWord := false) {
    global WD_ALERTS_NONE

    ; Validate files exist
    if !FileExist(templatePath)
        throw Error("Template file not found: " . templatePath)

    if !FileExist(dataSourcePath)
        throw Error("Data source file not found: " . dataSourcePath)

    ; Create Word application
    wordApp := ComObject("Word.Application")
    wordApp.Visible := showWord
    wordApp.DisplayAlerts := WD_ALERTS_NONE

    try {
        ; Open template as read-only to prevent accidental changes
        doc := wordApp.Documents.Open(templatePath, false, true)

        ; Connect to data source
        OpenDataSource(wordApp, doc, dataSourcePath)

        ; Get record count
        recordCount := GetRecordCount(doc)

        return {wordApp: wordApp, doc: doc, recordCount: recordCount}

    } catch as err {
        ; Cleanup on failure
        try {
            wordApp.Quit()
        }
        throw Error("Failed to initialize Word session: " . err.Message)
    }
}

/**
 * Cleans up Word session by closing document and quitting application
 * @param session - Object returned by InitializeWordSession
 */
CleanupWordSession(session) {
    global WD_DO_NOT_SAVE_CHANGES

    try {
        if session.HasProp("doc") && session.doc != "" {
            session.doc.Close(WD_DO_NOT_SAVE_CHANGES)
        }
    }

    try {
        if session.HasProp("wordApp") && session.wordApp != "" {
            session.wordApp.Quit()
        }
    }
}

; ========================================
; Phase 3: Document Merge Functions
; ========================================

/**
 * Sanitizes a filename by removing or replacing invalid characters
 * @param filename - Raw filename string to sanitize
 * @param replacement - Character to use for replacements (default: "_")
 * @returns Sanitized filename safe for Windows file system
 */
SanitizeFilename(filename, replacement := "_") {
    ; Return default if empty
    if !filename || filename = ""
        return "Document"

    ; Remove or replace invalid Windows filename characters: < > : " / \ | ? *
    sanitized := filename

    ; Replace each invalid character
    invalidChars := ['<', '>', ':', '"', '/', '\', '|', '?', '*']
    for char in invalidChars {
        sanitized := StrReplace(sanitized, char, replacement)
    }

    ; Remove leading/trailing spaces and periods (Windows doesn't allow)
    sanitized := Trim(sanitized)
    sanitized := RTrim(sanitized, ".")

    ; Replace multiple consecutive underscores with single underscore
    while InStr(sanitized, "__")
        sanitized := StrReplace(sanitized, "__", "_")

    ; Ensure filename isn't too long (Windows MAX_PATH consideration)
    if StrLen(sanitized) > 200
        sanitized := SubStr(sanitized, 1, 200)

    ; Return default if result is empty
    if !sanitized || sanitized = "" || sanitized = replacement
        return "Document"

    return sanitized
}

/**
 * Merges a single record to a new document and exports it as PDF
 * @param wordApp - Word Application COM object
 * @param templateDoc - Template document COM object
 * @param recordIndex - 1-based index of the record to merge
 * @param outputPath - Full path for the output PDF file (including extension)
 * @returns Boolean indicating success
 */
MergeRecordToDocument(wordApp, templateDoc, recordIndex, outputPath) {
    global WD_DO_NOT_SAVE_CHANGES
    global WD_SEND_TO_NEW_DOCUMENT, WD_STATISTIC_PAGES

    mergedDoc := ""

    try {
        ; Get template page count before merge (to avoid blank trailing pages in PDF)
        templatePageCount := templateDoc.ComputeStatistics(WD_STATISTIC_PAGES)

        ; Set the active record
        SetActiveRecord(templateDoc, recordIndex)

        ; Configure mail merge for single record output
        mailMerge := templateDoc.MailMerge
        mailMerge.Destination := WD_SEND_TO_NEW_DOCUMENT
        mailMerge.DataSource.FirstRecord := recordIndex
        mailMerge.DataSource.LastRecord := recordIndex

        ; Execute the merge - creates a new document with merged content
        mailMerge.Execute(false)

        ; Get the newly created merged document (it becomes the active document)
        mergedDoc := wordApp.ActiveDocument

        ; Export as PDF - limit to template page count to avoid blank trailing pages
        mergedDoc.ExportAsFixedFormat(
            outputPath,                      ; OutputFileName
            WD_EXPORT_FORMAT_PDF,           ; wdExportFormatPDF
            false,                           ; OpenAfterExport
            WD_EXPORT_OPTIMIZE_FOR_PRINT,   ; wdExportOptimizeForPrint
            WD_EXPORT_FROM_TO,              ; Export page range
            1,                               ; From: page 1
            templatePageCount                ; To: template page count
        )

        ; Close the merged document without saving
        mergedDoc.Close(WD_DO_NOT_SAVE_CHANGES)
        mergedDoc := ""

        return true

    } catch as err {
        ; Ensure merged doc is closed on error
        try {
            if mergedDoc != ""
                mergedDoc.Close(WD_DO_NOT_SAVE_CHANGES)
        }
        throw Error("Failed to merge record " . recordIndex . ": " . err.Message)
    }
}

/**
 * Main document merge function - creates individual files for each record
 * @param session - Word session object from InitializeWordSession
 * @param options - Object with: outputFolder, filenameField, prefix
 * @returns Object with success count, failure count, and errors array
 */
MergeToDocuments(session, options) {
    global IsProcessing

    results := {
        success: 0,
        failed: 0,
        errors: [],
        outputFolder: options.outputFolder
    }

    ; Ensure output folder exists
    if !DirExist(options.outputFolder) {
        try {
            DirCreate(options.outputFolder)
        } catch as err {
            throw Error("Failed to create output folder: " . err.Message)
        }
    }

    extension := ".pdf"

    ; Get total record count for progress
    totalRecords := session.recordCount

    UpdateStatus("Merging " . totalRecords . " documents...", "processing", 0)

    ; Process each record
    Loop totalRecords {
        ; Check for cancellation
        if !IsProcessing {
            results.errors.Push("Merge cancelled by user at record " . A_Index)
            break
        }

        recordIndex := A_Index

        try {
            ; Get the filename from the selected field
            SetActiveRecord(session.doc, recordIndex)
            rawFilename := GetFieldValue(session.doc, options.filenameField)

            ; Sanitize and build the full filename
            cleanFilename := SanitizeFilename(rawFilename)

            ; Add prefix if specified
            if options.prefix && options.prefix != ""
                cleanFilename := options.prefix . cleanFilename

            ; Build full output path
            outputPath := options.outputFolder . "\" . cleanFilename . extension

            ; Handle duplicate filenames by adding a number
            if FileExist(outputPath) {
                counter := 1
                Loop {
                    outputPath := options.outputFolder . "\" . cleanFilename . "_" . counter . extension
                    if !FileExist(outputPath)
                        break
                    counter++
                    if counter > 999 {
                        throw Error("Too many duplicate filenames")
                    }
                }
            }

            ; Update progress
            progress := Round((recordIndex / totalRecords) * 100)
            UpdateStatus("Processing record " . recordIndex . " of " . totalRecords . "...", "processing", progress)
            UpdateDetailedProgress("Creating: " . cleanFilename . extension)

            ; Merge this record
            MergeRecordToDocument(session.wordApp, session.doc, recordIndex, outputPath)

            results.success++

            ; Small delay to keep GUI responsive
            Sleep(10)

        } catch as err {
            results.failed++
            results.errors.Push("Record " . recordIndex . ": " . err.Message)
        }
    }

    return results
}

; ========================================
; Phase 4: Email Merge Functions
; ========================================

/**
 * Creates an Outlook email draft
 * @param toEmail - Email address to send to
 * @param subject - Email subject line
 * @param body - Email body content
 * @param options - Object with: bodyFormat ("html"/"plain"), cc, bcc, attachmentPaths (array)
 * @returns Boolean indicating success
 */
CreateOutlookDraft(toEmail, subject, body, options := {}) {
    global OL_MAIL_ITEM, OL_BODY_FORMAT_HTML, OL_BODY_FORMAT_PLAIN

    outlook := ""
    mail := ""

    try {
        ; Create or get existing Outlook application
        try {
            outlook := ComObjActive("Outlook.Application")
        } catch {
            outlook := ComObject("Outlook.Application")
        }

        ; Create new mail item
        mail := outlook.CreateItem(OL_MAIL_ITEM)

        ; Set recipient
        mail.To := toEmail

        ; Set subject
        mail.Subject := subject

        ; Set CC if provided
        if options.HasProp("cc") && options.cc != "" && options.cc != "(None)"
            mail.CC := options.cc

        ; Set BCC if provided
        if options.HasProp("bcc") && options.bcc != "" && options.bcc != "(None)"
            mail.BCC := options.bcc

        ; Set body content
        bodyFormat := options.HasProp("bodyFormat") ? options.bodyFormat : "plain"
        if bodyFormat = "html" {
            mail.HTMLBody := body
        } else {
            mail.Body := body
        }

        ; Add attachments if provided (supports array of paths)
        if options.HasProp("attachmentPaths") && options.attachmentPaths.Length > 0 {
            for attachPath in options.attachmentPaths {
                if attachPath != "" && FileExist(attachPath) {
                    mail.Attachments.Add(attachPath)
                }
            }
        }

        ; Display as draft (NOT Send)
        mail.Display()

        return true

    } catch as err {
        throw Error("Failed to create Outlook draft: " . err.Message)
    }
}

/**
 * Converts a Word document to HTML string for email body
 * @param wordApp - Word Application COM object
 * @param templateDoc - Template document COM object
 * @param recordIndex - 1-based index of the record to merge
 * @returns HTML string of the merged content
 */
MergeRecordToHTML(wordApp, templateDoc, recordIndex) {
    global WD_FORMAT_FILTERED_HTML, WD_DO_NOT_SAVE_CHANGES, WD_SEND_TO_NEW_DOCUMENT

    mergedDoc := ""
    tempFile := A_Temp . "\MergeTool_temp_" . A_TickCount . ".html"

    try {
        ; Set the active record
        SetActiveRecord(templateDoc, recordIndex)

        ; Configure mail merge for single record output
        mailMerge := templateDoc.MailMerge
        mailMerge.Destination := WD_SEND_TO_NEW_DOCUMENT
        mailMerge.DataSource.FirstRecord := recordIndex
        mailMerge.DataSource.LastRecord := recordIndex

        ; Execute the merge
        mailMerge.Execute(false)

        ; Get the newly created merged document
        mergedDoc := wordApp.ActiveDocument

        ; Save as Filtered HTML (cleaner output, better for email)
        ; Filtered HTML removes Office-specific tags and produces web-optimized HTML
        mergedDoc.SaveAs2(tempFile, WD_FORMAT_FILTERED_HTML)

        ; Close the merged document
        mergedDoc.Close(WD_DO_NOT_SAVE_CHANGES)
        mergedDoc := ""

        ; Read the HTML content
        htmlContent := FileRead(tempFile)

        ; Delete temp file
        try {
            FileDelete(tempFile)
        }

        ; Clean up any associated files folder (Word creates _files folder for HTML)
        filesFolder := StrReplace(tempFile, ".html", "_files")
        if DirExist(filesFolder) {
            try {
                DirDelete(filesFolder, true)
            }
        }

        ; Extract body content and process for email
        htmlContent := ProcessHTMLForEmail(htmlContent, filesFolder)

        return htmlContent

    } catch as err {
        ; Cleanup on error
        try {
            if mergedDoc != ""
                mergedDoc.Close(WD_DO_NOT_SAVE_CHANGES)
        }
        try {
            if FileExist(tempFile)
                FileDelete(tempFile)
        }
        throw Error("Failed to convert record " . recordIndex . " to HTML: " . err.Message)
    }
}

/**
 * Processes HTML content for email compatibility
 * Extracts body content and handles embedded images
 * @param htmlContent - Full HTML document string
 * @param filesFolder - Path to associated _files folder (if exists)
 * @returns Processed HTML suitable for email body
 */
ProcessHTMLForEmail(htmlContent, filesFolder) {
    ; Extract content between <body> tags, preserving all formatting
    if RegExMatch(htmlContent, "is)<body[^>]*>(.*?)</body>", &match) {
        bodyContent := match[1]
    } else {
        ; If no body tags found, use the entire content
        bodyContent := htmlContent
    }

    ; Convert image references to embedded base64 if images exist in _files folder
    ; This ensures images are displayed in emails even after temp files are deleted
    if DirExist(filesFolder) {
        ; Find all image references in the HTML
        bodyContent := RegExReplace(bodyContent, "i)src=[`"']([^`"']+)[`"']", ConvertImageToBase64.Bind(filesFolder))
    }

    ; Remove any remaining Office-specific metadata or XML islands
    bodyContent := RegExReplace(bodyContent, "is)<!\[if[^\]]*\]>.*?<!\[endif\]>", "")

    ; Clean up excessive whitespace while preserving intentional line breaks
    bodyContent := RegExReplace(bodyContent, "m)^\s+$", "")

    return bodyContent
}

/**
 * Callback function to convert image src to base64 embedded data
 * @param match - RegExMatch object containing the image src
 * @returns Replacement string with base64 data URI
 */
ConvertImageToBase64(filesFolder, match) {
    imagePath := match[1]

    ; Check if this is a local file reference (not already a data URI or external URL)
    if !RegExMatch(imagePath, "i)^(https?:|data:)") {
        ; Construct full path to the image file
        fullPath := filesFolder "\" imagePath

        if FileExist(fullPath) {
            try {
                ; Read image file as binary
                imageData := FileRead(fullPath, "RAW")

                ; Determine MIME type from extension
                if RegExMatch(imagePath, "i)\.(png)$")
                    mimeType := "image/png"
                else if RegExMatch(imagePath, "i)\.(jpe?g)$")
                    mimeType := "image/jpeg"
                else if RegExMatch(imagePath, "i)\.(gif)$")
                    mimeType := "image/gif"
                else if RegExMatch(imagePath, "i)\.(bmp)$")
                    mimeType := "image/bmp"
                else
                    mimeType := "image/png"  ; Default

                ; Convert to base64
                base64Data := Base64Encode(imageData)

                ; Return data URI
                return 'src="data:' mimeType ';base64,' base64Data '"'
            }
        }
    }

    ; If conversion failed or not applicable, return original
    return match[0]
}

/**
 * Encodes binary data to base64 string
 * @param data - Binary data to encode
 * @returns Base64 encoded string
 */
Base64Encode(data) {
    ; Get data size
    dataSize := data.Size

    ; Calculate required buffer size for base64 (4/3 of original, rounded up)
    base64Size := Ceil(dataSize * 4 / 3) + 4

    ; Create buffer for base64 output
    base64Buffer := Buffer(base64Size)

    ; Use Windows API to convert to base64
    DllCall("Crypt32\CryptBinaryToStringA",
        "Ptr", data.Ptr,
        "UInt", dataSize,
        "UInt", 0x40000001,  ; CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF
        "Ptr", base64Buffer.Ptr,
        "UInt*", &base64Size)

    ; Convert buffer to string
    return StrGet(base64Buffer.Ptr, base64Size, "CP0")
}

/**
 * Creates a temporary attachment file from merged document
 * @param wordApp - Word Application COM object
 * @param templateDoc - Template document COM object
 * @param recordIndex - 1-based index of the record to merge
 * @param format - "word" or "pdf"
 * @param filename - Base filename (without extension) for the attachment
 * @returns Full path to the created temporary file
 */
CreateTempAttachment(wordApp, templateDoc, recordIndex, format, filename) {
    global WD_FORMAT_DOCUMENT_DEFAULT, WD_FORMAT_PDF, WD_DO_NOT_SAVE_CHANGES
    global WD_SEND_TO_NEW_DOCUMENT, WD_STATISTIC_PAGES

    extension := (format = "pdf") ? ".pdf" : ".docx"
    tempPath := A_Temp . "\MergeTool_" . SanitizeFilename(filename) . "_" . A_TickCount . extension
    mergedDoc := ""

    try {
        ; Get template page count before merge (to avoid blank trailing pages in PDF)
        templatePageCount := templateDoc.ComputeStatistics(WD_STATISTIC_PAGES)

        ; Set the active record
        SetActiveRecord(templateDoc, recordIndex)

        ; Configure mail merge for single record output
        mailMerge := templateDoc.MailMerge
        mailMerge.Destination := WD_SEND_TO_NEW_DOCUMENT
        mailMerge.DataSource.FirstRecord := recordIndex
        mailMerge.DataSource.LastRecord := recordIndex

        ; Execute the merge
        mailMerge.Execute(false)

        ; Get the merged document
        mergedDoc := wordApp.ActiveDocument

        ; Save to temp location using appropriate method
        if (format = "pdf") {
            ; Use ExportAsFixedFormat for PDF - limit to template page count to avoid blank trailing pages
            mergedDoc.ExportAsFixedFormat(
                tempPath,                        ; OutputFileName
                WD_EXPORT_FORMAT_PDF,           ; wdExportFormatPDF
                false,                           ; OpenAfterExport
                WD_EXPORT_OPTIMIZE_FOR_PRINT,   ; wdExportOptimizeForPrint
                WD_EXPORT_FROM_TO,              ; Export page range
                1,                               ; From: page 1
                templatePageCount                ; To: template page count
            )
        } else {
            ; Use SaveAs2 for Word documents
            mergedDoc.SaveAs2(tempPath, WD_FORMAT_DOCUMENT_DEFAULT)
        }

        ; Close the merged document
        mergedDoc.Close(WD_DO_NOT_SAVE_CHANGES)
        mergedDoc := ""

        return tempPath

    } catch as err {
        try {
            if mergedDoc != ""
                mergedDoc.Close(WD_DO_NOT_SAVE_CHANGES)
        }
        throw Error("Failed to create attachment for record " . recordIndex . ": " . err.Message)
    }
}

/**
 * Main email merge function - creates Outlook drafts for each record
 * @param session - Word session object from InitializeWordSession
 * @param options - Object with: emailField, subjectField/staticSubject, ccField, bccField, attachField1/2/3, commonAttachment
 * @returns Object with success count, failure count, and errors array
 */
MergeToEmail(session, options) {
    global IsProcessing

    results := {
        success: 0,
        failed: 0,
        errors: []
    }

    totalRecords := session.recordCount

    UpdateStatus("Creating " . totalRecords . " email drafts...", "processing", 0)

    ; Process each record
    Loop totalRecords {
        ; Check for cancellation
        if !IsProcessing {
            results.errors.Push("Email merge cancelled by user at record " . A_Index)
            break
        }

        recordIndex := A_Index

        try {
            ; Set the active record to read field values
            SetActiveRecord(session.doc, recordIndex)

            ; Get email address
            toEmail := GetFieldValue(session.doc, options.emailField)

            ; Validate email is not empty
            if !toEmail || toEmail = "" {
                throw Error("Empty email address")
            }

            ; Get subject from field
            subject := ""
            if options.HasProp("subjectField") && options.subjectField != "" {
                subject := GetFieldValue(session.doc, options.subjectField)
            }

            ; Get CC (collect from all 3 field dropdowns + static)
            ccValues := []
            if options.HasProp("ccField") && options.ccField != "" && options.ccField != "(None)" {
                try {
                    val := GetFieldValue(session.doc, options.ccField)
                    if val != ""
                        ccValues.Push(val)
                }
            }
            if options.HasProp("ccField2") && options.ccField2 != "" && options.ccField2 != "(None)" {
                try {
                    val := GetFieldValue(session.doc, options.ccField2)
                    if val != ""
                        ccValues.Push(val)
                }
            }
            if options.HasProp("ccField3") && options.ccField3 != "" && options.ccField3 != "(None)" {
                try {
                    val := GetFieldValue(session.doc, options.ccField3)
                    if val != ""
                        ccValues.Push(val)
                }
            }
            ccValue := JoinArray(ccValues, "; ")

            ; Get BCC (collect from all 3 field dropdowns + static)
            bccValues := []
            if options.HasProp("bccField") && options.bccField != "" && options.bccField != "(None)" {
                try {
                    val := GetFieldValue(session.doc, options.bccField)
                    if val != ""
                        bccValues.Push(val)
                }
            }
            if options.HasProp("bccField2") && options.bccField2 != "" && options.bccField2 != "(None)" {
                try {
                    val := GetFieldValue(session.doc, options.bccField2)
                    if val != ""
                        bccValues.Push(val)
                }
            }
            if options.HasProp("bccField3") && options.bccField3 != "" && options.bccField3 != "(None)" {
                try {
                    val := GetFieldValue(session.doc, options.bccField3)
                    if val != ""
                        bccValues.Push(val)
                }
            }
            bccValue := JoinArray(bccValues, "; ")

            ; Update progress
            progress := Round((recordIndex / totalRecords) * 100)
            UpdateStatus("Creating draft " . recordIndex . " of " . totalRecords . "...", "processing", progress)
            UpdateDetailedProgress("To: " . toEmail)

            ; Collect per-person attachments from selected columns
            attachmentPaths := []

            ; Check attachment field 1
            if options.HasProp("attachField1") && options.attachField1 != "(None)" && options.attachField1 != "" {
                try {
                    filePath := GetFieldValue(session.doc, options.attachField1)
                    if filePath != "" {
                        attachmentPaths.Push(filePath)
                    }
                }
            }

            ; Check attachment field 2
            if options.HasProp("attachField2") && options.attachField2 != "(None)" && options.attachField2 != "" {
                try {
                    filePath := GetFieldValue(session.doc, options.attachField2)
                    if filePath != "" {
                        attachmentPaths.Push(filePath)
                    }
                }
            }

            ; Check attachment field 3
            if options.HasProp("attachField3") && options.attachField3 != "(None)" && options.attachField3 != "" {
                try {
                    filePath := GetFieldValue(session.doc, options.attachField3)
                    if filePath != "" {
                        attachmentPaths.Push(filePath)
                    }
                }
            }

            ; Add common attachment if specified
            if options.HasProp("commonAttachment") && options.commonAttachment != "" {
                attachmentPaths.Push(options.commonAttachment)
            }

            ; Email body is always the merged template HTML
            body := MergeRecordToHTML(session.wordApp, session.doc, recordIndex)

            ; Prepare draft options
            draftOptions := {
                cc: ccValue,
                bcc: bccValue,
                bodyFormat: "html",
                attachmentPaths: attachmentPaths
            }

            ; Create the Outlook draft
            CreateOutlookDraft(toEmail, subject, body, draftOptions)

            results.success++

            ; Small delay to keep GUI responsive and allow Outlook to process
            Sleep(100)

        } catch as err {
            results.failed++
            results.errors.Push("Record " . recordIndex . ": " . err.Message)
        }
    }

    return results
}

/**
 * Main email merge function - sends emails directly to Outbox (no preview)
 * @param session - Word session object from InitializeWordSession
 * @param options - Object with: emailField, subjectField/staticSubject, ccField, bccField, attachField1/2/3, commonAttachment, sendFromAccount
 * @returns Object with success count, failure count, and errors array
 */
MergeToEmailDirect(session, options) {
    global IsProcessing, OL_MAIL_ITEM

    results := {
        success: 0,
        failed: 0,
        errors: []
    }

    totalRecords := session.recordCount

    ; Get or create Outlook
    outlook := ""
    try {
        try {
            outlook := ComObjActive("Outlook.Application")
        } catch {
            outlook := ComObject("Outlook.Application")
        }
    } catch as err {
        throw Error("Failed to connect to Outlook: " . err.Message)
    }

    ; Find the selected send-from account (if not default)
    sendFromAccount := ""
    if options.HasProp("sendFromAccount") && options.sendFromAccount != "(Default Account)" && options.sendFromAccount != "" {
        try {
            namespace := outlook.GetNamespace("MAPI")
            accountCount := namespace.Accounts.Count

            ; Extract just the email address from the dropdown selection (format: "email@example.com (Display Name)")
            selectedAccountEmail := options.sendFromAccount
            if InStr(selectedAccountEmail, " (") {
                selectedAccountEmail := SubStr(selectedAccountEmail, 1, InStr(selectedAccountEmail, " (") - 1)
            }

            ; Find matching account
            Loop accountCount {
                account := namespace.Accounts.Item(A_Index)
                if (account.SmtpAddress = selectedAccountEmail) {
                    sendFromAccount := account
                    break
                }
            }
        } catch as err {
            ; If we can't find the account, continue with default
            results.errors.Push("Warning: Could not set send-from account, using default: " . err.Message)
        }
    }

    UpdateStatus("Sending " . totalRecords . " emails...", "processing", 0)

    ; Process each record
    Loop totalRecords {
        ; Check for cancellation
        if !IsProcessing {
            results.errors.Push("Email merge cancelled by user at record " . A_Index)
            break
        }

        recordIndex := A_Index

        try {
            ; Set the active record to read field values
            SetActiveRecord(session.doc, recordIndex)

            ; Get email address
            toEmail := GetFieldValue(session.doc, options.emailField)

            ; Validate email is not empty
            if !toEmail || toEmail = "" {
                throw Error("Empty email address")
            }

            ; Get subject from field
            subject := ""
            if options.HasProp("subjectField") && options.subjectField != "" {
                subject := GetFieldValue(session.doc, options.subjectField)
            }

            ; Get CC (collect from all 3 field dropdowns + static)
            ccValues := []
            if options.HasProp("ccField") && options.ccField != "" && options.ccField != "(None)" {
                try {
                    val := GetFieldValue(session.doc, options.ccField)
                    if val != ""
                        ccValues.Push(val)
                }
            }
            if options.HasProp("ccField2") && options.ccField2 != "" && options.ccField2 != "(None)" {
                try {
                    val := GetFieldValue(session.doc, options.ccField2)
                    if val != ""
                        ccValues.Push(val)
                }
            }
            if options.HasProp("ccField3") && options.ccField3 != "" && options.ccField3 != "(None)" {
                try {
                    val := GetFieldValue(session.doc, options.ccField3)
                    if val != ""
                        ccValues.Push(val)
                }
            }
            ccValue := JoinArray(ccValues, "; ")

            ; Get BCC (collect from all 3 field dropdowns + static)
            bccValues := []
            if options.HasProp("bccField") && options.bccField != "" && options.bccField != "(None)" {
                try {
                    val := GetFieldValue(session.doc, options.bccField)
                    if val != ""
                        bccValues.Push(val)
                }
            }
            if options.HasProp("bccField2") && options.bccField2 != "" && options.bccField2 != "(None)" {
                try {
                    val := GetFieldValue(session.doc, options.bccField2)
                    if val != ""
                        bccValues.Push(val)
                }
            }
            if options.HasProp("bccField3") && options.bccField3 != "" && options.bccField3 != "(None)" {
                try {
                    val := GetFieldValue(session.doc, options.bccField3)
                    if val != ""
                        bccValues.Push(val)
                }
            }
            bccValue := JoinArray(bccValues, "; ")

            ; Update progress
            progress := Round((recordIndex / totalRecords) * 100)
            UpdateStatus("Sending email " . recordIndex . " of " . totalRecords . "...", "processing", progress)
            UpdateDetailedProgress("To: " . toEmail)

            ; Collect per-person attachments from selected columns
            attachmentPaths := []

            if options.HasProp("attachField1") && options.attachField1 != "(None)" && options.attachField1 != "" {
                try {
                    filePath := GetFieldValue(session.doc, options.attachField1)
                    if filePath != ""
                        attachmentPaths.Push(filePath)
                }
            }

            if options.HasProp("attachField2") && options.attachField2 != "(None)" && options.attachField2 != "" {
                try {
                    filePath := GetFieldValue(session.doc, options.attachField2)
                    if filePath != ""
                        attachmentPaths.Push(filePath)
                }
            }

            if options.HasProp("attachField3") && options.attachField3 != "(None)" && options.attachField3 != "" {
                try {
                    filePath := GetFieldValue(session.doc, options.attachField3)
                    if filePath != ""
                        attachmentPaths.Push(filePath)
                }
            }

            ; Add common attachment if specified
            if options.HasProp("commonAttachment") && options.commonAttachment != "" {
                attachmentPaths.Push(options.commonAttachment)
            }

            ; Email body is the merged template HTML
            body := MergeRecordToHTML(session.wordApp, session.doc, recordIndex)

            ; Create and send the email
            mail := outlook.CreateItem(OL_MAIL_ITEM)
            mail.To := toEmail
            mail.Subject := subject

            if ccValue != ""
                mail.CC := ccValue
            if bccValue != ""
                mail.BCC := bccValue

            mail.HTMLBody := body

            ; Add attachments
            for attachPath in attachmentPaths {
                if attachPath != "" && FileExist(attachPath) {
                    mail.Attachments.Add(attachPath)
                }
            }

            ; Set the send-from account if specified
            if (sendFromAccount != "") {
                mail.SendUsingAccount := sendFromAccount
            }

            ; Send to Outbox
            mail.Send()

            results.success++

            ; Delay between sends to avoid triggering spam/rate-limit blocks
            Sleep(Round(options.delaySeconds * 1000))

        } catch as err {
            results.failed++
            results.errors.Push("Record " . recordIndex . ": " . err.Message)
        }
    }

    return results
}

; ========================================
; Phase 5: Printer Merge Functions
; ========================================

/**
 * Main printer merge function - sends merged documents to printer
 * @param session - Word session object from InitializeWordSession
 * @param options - Object with: printerName, copies, collate
 * @returns Object with success status and any errors
 */
MergeToPrinter(session, options) {
    global IsProcessing, WD_SEND_TO_NEW_DOCUMENT, WD_DO_NOT_SAVE_CHANGES

    mergedDoc := ""

    results := {
        success: false,
        recordCount: session.recordCount,
        copies: options.copies,
        errors: []
    }

    try {
        UpdateStatus("Preparing print job for " . session.recordCount . " records...", "processing", 20)

        ; Set the active printer
        if options.HasProp("printerName") && options.printerName != "" {
            session.wordApp.ActivePrinter := options.printerName
            UpdateDetailedProgress("Printer: " . options.printerName)
        }

        UpdateStatus("Configuring mail merge for printing...", "processing", 40)

        ; Merge every record into a single document rather than straight to the
        ; printer. Copies and collation are properties of a print job, not of the
        ; merge, so they can only be applied by calling PrintOut on the result.
        mailMerge := session.doc.MailMerge
        mailMerge.Destination := WD_SEND_TO_NEW_DOCUMENT

        ; Set record range (all records)
        mailMerge.DataSource.FirstRecord := 1
        mailMerge.DataSource.LastRecord := -16  ; wdDefaultLastRecord - all records

        session.wordApp.Options.PrintBackground := false  ; Print synchronously

        UpdateStatus("Merging " . session.recordCount . " records...", "processing", 55)

        if !IsProcessing {
            results.errors.Push("Print merge cancelled by user")
            return results
        }

        mailMerge.Execute(false)  ; false = don't pause
        mergedDoc := session.wordApp.ActiveDocument

        UpdateStatus("Sending " . session.recordCount . " documents to printer...", "processing", 75)

        ; PrintOut applies copies and collation to the job as a whole. With
        ; Collate on, each copy prints as a complete set; with it off, each page
        ; prints its copies together.
        mergedDoc.PrintOut(
            false,              ; Background - wait for the job to spool
            ,                   ; Append
            ,                   ; Range - whole document
            ,                   ; OutputFileName
            ,                   ; From
            ,                   ; To
            ,                   ; Item
            options.copies,     ; Copies
            ,                   ; Pages
            ,                   ; PageType
            ,                   ; PrintToFile
            options.collate     ; Collate
        )

        results.success := true
        UpdateStatus("Print job sent successfully", "success", 100)

    } catch as err {
        results.errors.Push(err.Message)
        throw Error("Print merge failed: " . err.Message)
    } finally {
        ; Discard the intermediate merged document either way
        try {
            if mergedDoc != ""
                mergedDoc.Close(WD_DO_NOT_SAVE_CHANGES)
        }
    }

    return results
}

/**
 * Gets the default printer name from Windows
 * @returns String name of the default printer
 */
GetDefaultPrinter() {
    try {
        ; Read default printer from registry
        defaultPrinter := RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows", "Device")
        ; Format is "PrinterName,winspool,Ne00:" - extract just the name
        if InStr(defaultPrinter, ",") {
            defaultPrinter := SubStr(defaultPrinter, 1, InStr(defaultPrinter, ",") - 1)
        }
        return defaultPrinter
    } catch {
        return ""
    }
}

; ========================================
; Populate Field Dropdowns
; ========================================
PopulateFieldDropdowns() {
    global FieldNames
    global EmailFieldDropdown, SubjectFieldDropdown
    global CCFieldDropdown, CCField2Dropdown, CCField3Dropdown
    global BCCFieldDropdown, BCCField2Dropdown, BCCField3Dropdown
    global FilenameFieldDropdown
    global AttachmentField1DD, AttachmentField2DD, AttachmentField3DD

    ; Clear and populate email field dropdown
    EmailFieldDropdown.Delete()
    for field in FieldNames {
        EmailFieldDropdown.Add([field])
    }
    if FieldNames.Length > 0
        EmailFieldDropdown.Choose(1)

    ; Subject field dropdown
    SubjectFieldDropdown.Delete()
    for field in FieldNames {
        SubjectFieldDropdown.Add([field])
    }
    if FieldNames.Length > 0
        SubjectFieldDropdown.Choose(1)

    ; CC field dropdowns (with None option)
    CCFieldDropdown.Delete()
    CCFieldDropdown.Add(["(None)"])
    for field in FieldNames {
        CCFieldDropdown.Add([field])
    }
    CCFieldDropdown.Choose(1)

    CCField2Dropdown.Delete()
    CCField2Dropdown.Add(["(None)"])
    for field in FieldNames {
        CCField2Dropdown.Add([field])
    }
    CCField2Dropdown.Choose(1)

    CCField3Dropdown.Delete()
    CCField3Dropdown.Add(["(None)"])
    for field in FieldNames {
        CCField3Dropdown.Add([field])
    }
    CCField3Dropdown.Choose(1)

    ; BCC field dropdowns (with None option)
    BCCFieldDropdown.Delete()
    BCCFieldDropdown.Add(["(None)"])
    for field in FieldNames {
        BCCFieldDropdown.Add([field])
    }
    BCCFieldDropdown.Choose(1)

    BCCField2Dropdown.Delete()
    BCCField2Dropdown.Add(["(None)"])
    for field in FieldNames {
        BCCField2Dropdown.Add([field])
    }
    BCCField2Dropdown.Choose(1)

    BCCField3Dropdown.Delete()
    BCCField3Dropdown.Add(["(None)"])
    for field in FieldNames {
        BCCField3Dropdown.Add([field])
    }
    BCCField3Dropdown.Choose(1)

    ; Attachment field dropdowns (with None option)
    AttachmentField1DD.Delete()
    AttachmentField1DD.Add(["(None)"])
    for field in FieldNames {
        AttachmentField1DD.Add([field])
    }
    AttachmentField1DD.Choose(1)

    AttachmentField2DD.Delete()
    AttachmentField2DD.Add(["(None)"])
    for field in FieldNames {
        AttachmentField2DD.Add([field])
    }
    AttachmentField2DD.Choose(1)

    AttachmentField3DD.Delete()
    AttachmentField3DD.Add(["(None)"])
    for field in FieldNames {
        AttachmentField3DD.Add([field])
    }
    AttachmentField3DD.Choose(1)

    ; Filename field dropdown
    FilenameFieldDropdown.Delete()
    for field in FieldNames {
        FilenameFieldDropdown.Add([field])
    }
    if FieldNames.Length > 0
        FilenameFieldDropdown.Choose(1)
}

; ========================================
; Enumerate Outlook Accounts
; ========================================
EnumerateOutlookAccounts() {
    global SendFromAccountDropdown

    accounts := []

    try {
        ; Get or create Outlook application
        outlook := ""
        try {
            outlook := ComObjActive("Outlook.Application")
        } catch {
            outlook := ComObject("Outlook.Application")
        }

        ; Get namespace and enumerate accounts
        namespace := outlook.GetNamespace("MAPI")
        accountCount := namespace.Accounts.Count

        Loop accountCount {
            account := namespace.Accounts.Item(A_Index)
            ; Store both display name and email address for clarity
            displayName := account.DisplayName
            smtpAddress := account.SmtpAddress

            ; Format: "email@example.com (Display Name)" or just email if same
            if (displayName != smtpAddress && displayName != "") {
                accounts.Push(smtpAddress . " (" . displayName . ")")
            } else {
                accounts.Push(smtpAddress)
            }
        }

    } catch as err {
        ; Silently fail - will just show default account option
    }

    ; Update dropdown
    SendFromAccountDropdown.Delete()
    SendFromAccountDropdown.Add(["(Default Account)"])

    if accounts.Length > 0 {
        for account in accounts {
            SendFromAccountDropdown.Add([account])
        }
    }

    ; Try to select the last saved account
    global LastSavedAccount
    selectedAccount := false
    if IsSet(LastSavedAccount) && LastSavedAccount != "" {
        ; Build complete list including default account
        allAccounts := ["(Default Account)"]
        if accounts.Length > 0 {
            allAccounts.Push(accounts*)
        }

        ; Find the saved account in the list
        Loop allAccounts.Length {
            if allAccounts[A_Index] = LastSavedAccount {
                SendFromAccountDropdown.Choose(A_Index)
                selectedAccount := true
                break
            }
        }
    }

    ; If saved account not found or no saved account, select first (Default Account)
    if !selectedAccount {
        SendFromAccountDropdown.Choose(1)
    }
}

; ========================================
; Enumerate System Printers
; ========================================
EnumeratePrinters() {
    global PrinterDropdown, LastSavedPrinter

    printers := []

    ; Read printers from registry
    try {
        Loop Reg, "HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Devices", "V" {
            printers.Push(A_LoopRegName)
        }
    }

    ; Update dropdown
    PrinterDropdown.Delete()
    if printers.Length > 0 {
        for printer in printers {
            PrinterDropdown.Add([printer])
        }

        ; Try to select the last saved printer
        selectedPrinter := false
        if IsSet(LastSavedPrinter) && LastSavedPrinter != "" {
            ; Find the saved printer in the list
            Loop printers.Length {
                if printers[A_Index] = LastSavedPrinter {
                    PrinterDropdown.Choose(A_Index)
                    selectedPrinter := true
                    break
                }
            }
        }

        ; If saved printer not found or no saved printer, select first
        if !selectedPrinter {
            PrinterDropdown.Choose(1)
        }
    } else {
        PrinterDropdown.Add(["(No printers found)"])
        PrinterDropdown.Choose(1)
    }
}

; ========================================
; Start Merge
; ========================================
StartMerge(*) {
    global CurrentMergeType, IsProcessing
    global SelectedTemplatePath, SelectedDataSourcePath, FieldNames
    global FilenameFieldDropdown
    global FilenamePrefixEdit, OutputFolderEdit
    global EmailFieldDropdown, SubjectFieldDropdown
    global CCFieldDropdown, BCCFieldDropdown
    global AttachmentField1DD, AttachmentField2DD, AttachmentField3DD
    global CommonAttachmentPath, SendFromAccountDropdown
    global DelaySecondsEdit
    global PrinterDropdown, CopiesEdit, CollateCheckbox

    ; Validate inputs
    if !ValidateInputs()
        return

    IsProcessing := true
    SetControlsEnabled(false)
    UpdateStatus("Starting " . CurrentMergeType . " merge...", "processing", 0)

    session := ""

    try {
        ; Initialize Word session
        UpdateStatus("Connecting to data source...", "processing", 10)
        session := InitializeWordSession(SelectedTemplatePath, SelectedDataSourcePath)

        ; Route to appropriate merge function based on type
        switch CurrentMergeType {
            case "documents":
                ; Build options for document merge
                options := {
                    outputFolder: OutputFolderEdit.Value,
                    filenameField: FilenameFieldDropdown.Text,
                    prefix: FilenamePrefixEdit.Value
                }

                ; Execute document merge
                results := MergeToDocuments(session, options)

                ; Show completion summary
                UpdateStatus("Merge complete: " . results.success . " succeeded, " . results.failed . " failed",
                    results.failed > 0 ? "error" : "success", 100)
                UpdateDetailedProgress("Output folder: " . results.outputFolder)

                ; Show detailed results
                summaryMsg := "Document Merge Complete`n`n"
                    . "Successfully created: " . results.success . " files`n"
                    . "Failed: " . results.failed . " files`n"
                    . "Output folder: " . results.outputFolder

                if results.errors.Length > 0 {
                    summaryMsg .= "`n`nErrors:`n"
                    Loop Min(10, results.errors.Length) {
                        summaryMsg .= "• " . results.errors[A_Index] . "`n"
                    }
                    if results.errors.Length > 10
                        summaryMsg .= "... and " . (results.errors.Length - 10) . " more errors"
                }

                MsgBox(summaryMsg, "Merge Complete", results.failed > 0 ? "Icon!" : "Iconi")

                ; Offer to open output folder
                if results.success > 0 {
                    if MsgBox("Would you like to open the output folder?", "Open Folder", "YesNo") = "Yes" {
                        Run('explorer.exe "' . results.outputFolder . '"')
                    }
                }

            case "email":
                ; Build options for email merge
                options := {
                    emailField: EmailFieldDropdown.Text,
                    subjectField: SubjectFieldDropdown.Text,
                    ccField: CCFieldDropdown.Text,
                    ccField2: CCField2Dropdown.Text,
                    ccField3: CCField3Dropdown.Text,
                    bccField: BCCFieldDropdown.Text,
                    bccField2: BCCField2Dropdown.Text,
                    bccField3: BCCField3Dropdown.Text,
                    attachField1: AttachmentField1DD.Text,
                    attachField2: AttachmentField2DD.Text,
                    attachField3: AttachmentField3DD.Text,
                    commonAttachment: CommonAttachmentPath,
                    sendFromAccount: SendFromAccountDropdown.Text,
                    delaySeconds: Number(DelaySecondsEdit.Value)
                }

                ; Send directly to Outbox
                results := MergeToEmailDirect(session, options)

                ; Show completion summary
                UpdateStatus("Email merge complete: " . results.success . " sent, " . results.failed . " failed",
                    results.failed > 0 ? "error" : "success", 100)
                UpdateDetailedProgress("Emails have been placed in your Outlook Outbox")

                summaryMsg := "Email Merge Complete`n`n"
                    . "Successfully sent: " . results.success . " emails`n"
                    . "Failed: " . results.failed . " emails`n`n"
                    . "Emails have been placed in your Outlook Outbox."

                if results.errors.Length > 0 {
                    summaryMsg .= "`n`nErrors:`n"
                    Loop Min(10, results.errors.Length) {
                        summaryMsg .= "• " . results.errors[A_Index] . "`n"
                    }
                    if results.errors.Length > 10
                        summaryMsg .= "... and " . (results.errors.Length - 10) . " more errors"
                }

                MsgBox(summaryMsg, "Email Merge Complete", results.failed > 0 ? "Icon!" : "Iconi")

            case "print":
                ; Build options for print merge
                options := {
                    printerName: PrinterDropdown.Text,
                    copies: Integer(CopiesEdit.Value),
                    collate: CollateCheckbox.Value
                }

                ; Confirm before printing
                confirmMsg := "Print Merge Confirmation`n`n"
                    . "Printer: " . options.printerName . "`n"
                    . "Records to print: " . session.recordCount . "`n"
                    . "Copies per record: " . options.copies . "`n"
                    . "Total documents to print: " . (session.recordCount * options.copies) . "`n`n"
                    . "Do you want to proceed?"

                if MsgBox(confirmMsg, "Confirm Print", "YesNo Icon?") != "Yes" {
                    UpdateStatus("Print merge cancelled", "normal", 0)
                    return
                }

                ; Execute print merge
                results := MergeToPrinter(session, options)

                ; Show completion summary
                if results.success {
                    UpdateStatus("Print merge complete: " . results.recordCount . " documents sent to printer",
                        "success", 100)
                    UpdateDetailedProgress("Printer: " . options.printerName)

                    summaryMsg := "Print Merge Complete`n`n"
                        . "Successfully sent to printer: " . results.recordCount . " documents`n"
                        . "Copies printed: " . results.copies . "`n"
                        . "Printer: " . options.printerName . "`n`n"
                        . "Check your printer for the output."

                    MsgBox(summaryMsg, "Print Complete", "Iconi")
                } else {
                    errorMsg := "Print merge encountered errors:`n`n"
                    for err in results.errors {
                        errorMsg .= "• " . err . "`n"
                    }
                    MsgBox(errorMsg, "Print Error", "Icon!")
                }
        }

    } catch as err {
        UpdateStatus("Error: " . err.Message, "error", 0)
        MsgBox("Error during merge:`n`n" . err.Message, "Error", "IconX")

    } finally {
        if session != ""
            CleanupWordSession(session)
        IsProcessing := false
        SetControlsEnabled(true)
        if !InStr(StatusText.Value, "complete")
            UpdateStatus("Ready", "normal", 0)
    }
}

; ========================================
; Cancel Merge
; ========================================
CancelMerge(*) {
    global IsProcessing

    if IsProcessing {
        IsProcessing := false
        UpdateStatus("Merge cancelled by user", "error", 0)
        SetControlsEnabled(true)
    }
}

; ========================================
; Input Validation
; ========================================
ValidateInputs() {
    global SelectedTemplatePath, SelectedDataSourcePath, CurrentMergeType
    global FieldNames, OutputFolderEdit
    global EmailFieldDropdown, SubjectFieldDropdown
    global CommonAttachmentPath, PrinterDropdown

    ; Check template
    if !SelectedTemplatePath || SelectedTemplatePath = "" {
        MsgBox("Please select a Word template.", "Validation Error", "Icon!")
        UpdateStatus("Error: No template selected", "error")
        return false
    }

    if !FileExist(SelectedTemplatePath) {
        MsgBox("Template file not found.", "Validation Error", "Icon!")
        UpdateStatus("Error: Template file not found", "error")
        return false
    }

    ; Check data source
    if !SelectedDataSourcePath || SelectedDataSourcePath = "" {
        MsgBox("Please select an Excel data source.", "Validation Error", "Icon!")
        UpdateStatus("Error: No data source selected", "error")
        return false
    }

    if !FileExist(SelectedDataSourcePath) {
        MsgBox("Data source file not found.", "Validation Error", "Icon!")
        UpdateStatus("Error: Data source file not found", "error")
        return false
    }

    ; Check fields are loaded
    if FieldNames.Length = 0 {
        MsgBox("Please load fields from the data source first.", "Validation Error", "Icon!")
        UpdateStatus("Error: Fields not loaded", "error")
        return false
    }

    ; Type-specific validation
    if CurrentMergeType = "documents" {
        if OutputFolderEdit.Value = "" {
            MsgBox("Please select an output folder for documents.", "Validation Error", "Icon!")
            UpdateStatus("Error: No output folder selected", "error")
            return false
        }
    }

    if CurrentMergeType = "email" {
        ; Check email field is selected
        if EmailFieldDropdown.Text = "" || EmailFieldDropdown.Text = "(Load fields first)" {
            MsgBox("Please select an email field.", "Validation Error", "Icon!")
            UpdateStatus("Error: No email field selected", "error")
            return false
        }

        ; Check subject field is selected
        if SubjectFieldDropdown.Text = "" || SubjectFieldDropdown.Text = "(Load fields first)" {
            MsgBox("Please select a subject field.", "Validation Error", "Icon!")
            UpdateStatus("Error: No subject field selected", "error")
            return false
        }

        ; Validate common attachment exists if specified
        if CommonAttachmentPath != "" && !FileExist(CommonAttachmentPath) {
            MsgBox("Common attachment file not found: " . CommonAttachmentPath, "Validation Error", "Icon!")
            UpdateStatus("Error: Common attachment file not found", "error")
            return false
        }
    }

    if CurrentMergeType = "print" {
        ; Check printer is selected
        if PrinterDropdown.Text = "" || PrinterDropdown.Text = "(No printers found)" || PrinterDropdown.Text = "(Loading printers...)" {
            MsgBox("Please select a printer.", "Validation Error", "Icon!")
            UpdateStatus("Error: No printer selected", "error")
            return false
        }

        ; Validate copies is a valid number
        copies := CopiesEdit.Value
        if !IsInteger(copies) || copies < 1 {
            MsgBox("Please enter a valid number of copies (1 or more).", "Validation Error", "Icon!")
            UpdateStatus("Error: Invalid number of copies", "error")
            return false
        }
    }

    return true
}

; ========================================
; Status and Progress Functions
; ========================================
UpdateStatus(message, statusType := "normal", progress := -1) {
    global StatusText, ProgressBar1

    StatusText.Value := message

    ; Set color based on status type
    switch statusType {
        case "success":
            StatusText.SetFont("c" . SUCCESS_COLOR . " bold")
        case "error":
            StatusText.SetFont("c" . ERROR_COLOR . " bold")
        case "processing":
            StatusText.SetFont("c" . ACCENT_COLOR . " bold")
        default:
            StatusText.SetFont("c" . TEXT_COLOR)
    }

    ; Update progress bar if specified
    if progress >= 0 && progress <= 100 {
        ProgressBar1.Value := progress
    }
}

UpdateDetailedProgress(message) {
    global DetailedProgress
    DetailedProgress.Value := message
}

; ========================================
; Control Enable/Disable
; ========================================
SetControlsEnabled(enabled) {
    global StartMergeBtn, CancelBtn, LoadFieldsBtn
    global MergeTypeEmail, MergeTypeDocuments, MergeTypePrint

    StartMergeBtn.Enabled := enabled
    CancelBtn.Enabled := !enabled
    LoadFieldsBtn.Enabled := enabled
    MergeTypeEmail.Enabled := enabled
    MergeTypeDocuments.Enabled := enabled
    MergeTypePrint.Enabled := enabled
}

; ========================================
; Preference Management
; ========================================
SavePreference(key, value) {
    global CONFIG_FILE
    IniWrite(value, CONFIG_FILE, "Settings", key)
}

LoadPreference(key, default := "") {
    global CONFIG_FILE
    if FileExist(CONFIG_FILE) {
        return IniRead(CONFIG_FILE, "Settings", key, default)
    }
    return default
}

SavePrinterPreference(*) {
    global PrinterDropdown

    ; Only save if a valid printer is selected
    printerName := PrinterDropdown.Text
    if printerName != "" && printerName != "(No printers found)" && printerName != "(Loading printers...)" {
        SavePreference("LastPrinter", printerName)
    }
}

SaveAccountPreference(*) {
    global SendFromAccountDropdown

    ; Only save if a valid account is selected
    accountName := SendFromAccountDropdown.Text
    if accountName != "" {
        SavePreference("LastSendFromAccount", accountName)
    }
}

SaveDelayPreference(*) {
    global DelaySecondsEdit

    SavePreference("LastEmailDelaySeconds", DelaySecondsEdit.Value)
}

LoadPreferences() {
    global SelectedTemplatePath, SelectedDataSourcePath, CurrentMergeType
    global TemplatePathEdit, DataSourcePathEdit, OutputFolderEdit
    global MergeTypeEmail, MergeTypeDocuments, MergeTypePrint

    ; Load last template
    lastTemplate := LoadPreference("LastTemplate", "")
    if lastTemplate != "" && FileExist(lastTemplate) {
        SelectedTemplatePath := lastTemplate
        TemplatePathEdit.Value := lastTemplate
    }

    ; Load last data source
    lastDataSource := LoadPreference("LastDataSource", "")
    if lastDataSource != "" && FileExist(lastDataSource) {
        SelectedDataSourcePath := lastDataSource
        DataSourcePathEdit.Value := lastDataSource
    }

    ; Load last output folder
    lastOutputFolder := LoadPreference("LastOutputFolder", "")
    if lastOutputFolder != "" {
        OutputFolderEdit.Value := lastOutputFolder
    }

    ; Load last merge type
    lastMergeType := LoadPreference("LastMergeType", "email")
    CurrentMergeType := lastMergeType

    switch lastMergeType {
        case "email":
            MergeTypeEmail.Value := 1
        case "documents":
            MergeTypeDocuments.Value := 1
        case "print":
            MergeTypePrint.Value := 1
    }

    ; Load last printer
    global PrinterDropdown, LastSavedPrinter
    lastPrinter := LoadPreference("LastPrinter", "")
    if lastPrinter != "" {
        ; Store for use after EnumeratePrinters runs
        LastSavedPrinter := lastPrinter
    }

    ; Load last send from account
    global SendFromAccountDropdown, LastSavedAccount
    lastAccount := LoadPreference("LastSendFromAccount", "")
    if lastAccount != "" {
        ; Store for use after EnumerateOutlookAccounts runs
        LastSavedAccount := lastAccount
    }

    ; Load last email send delay (default 15 seconds)
    global DelaySecondsEdit
    DelaySecondsEdit.Value := LoadPreference("LastEmailDelaySeconds", 15)

    ShowPanel(CurrentMergeType)
}

