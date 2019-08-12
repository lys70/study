Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Security

' 어셈블리에 대한 일반 정보는 다음 특성 집합을 통해 
' 제어됩니다. 어셈블리와 관련된 정보를 수정하려면
' 이러한 특성 값을 변경하세요.

' 어셈블리 특성 값을 검토합니다.

<Assembly: AssemblyTitle("ExcelAddIn1")> 
<Assembly: AssemblyDescription("")> 
<Assembly: AssemblyCompany("")> 
<Assembly: AssemblyProduct("ExcelAddIn1")> 
<Assembly: AssemblyCopyright("Copyright ©  2019")> 
<Assembly: AssemblyTrademark("")> 

' ComVisible을 false로 설정하면 이 어셈블리의 형식이 COM 구성 요소에 
' 표시되지 않습니다.  COM에서 이 어셈블리의 형식에 액세스하려면 
' 해당 형식에 대해 ComVisible 특성을 true로 설정하세요.
<Assembly: ComVisible(False)>

'이 프로젝트가 COM에 노출되는 경우 다음 GUID는 typelib의 ID를 나타냅니다.
<Assembly: Guid("99e8faf9-9ec6-4b4b-a578-97be3b793899")> 

' 어셈블리의 버전 정보는 다음 네 가지 값으로 구성됩니다.
'
'      주 버전
'      부 버전 
'      빌드 번호
'      수정 버전
'
' 모든 값을 지정하거나 아래와 같이 '*'를 사용하여 빌드 번호 및 수정 번호가 자동으로 
' 지정되도록 할 수 있습니다.
' <Assembly: AssemblyVersion("1.0.*")> 

<Assembly: AssemblyVersion("1.0.0.0")> 
<Assembly: AssemblyFileVersion("1.0.0.0")> 

Friend Module DesignTimeConstants
    Public Const RibbonTypeSerializer As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Serialization.RibbonTypeCodeDomSerializer, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
    Public Const RibbonBaseTypeSerializer As String = "System.ComponentModel.Design.Serialization.TypeCodeDomSerializer, System.Design"
    Public Const RibbonDesigner As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Design.RibbonDesigner, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
End Module
