<OpenSpanDesignDocument Version="8.0.1026.0" Serializer="2.0" Culture="en-US">
  <ComponentInfo>
    <Type Value="OpenSpan.Automation.GlobalContainer" />
    <Assembly Value="OpenSpan.Automation" />
    <AssemblyReferences>
      <Assembly Value="mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" />
      <Assembly Value="System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" />
      <Assembly Value="System.Data, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" />
      <Assembly Value="System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a" />
      <Assembly Value="System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" />
      <Assembly Value="Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" />
      <Assembly Value="OpenSpan, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" />
      <Assembly Value="OpenSpan.Automation, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" />
      <Assembly Value="OpenSpan.Controls, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" />
      <Assembly Value="OpenSpan.Office, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" />
      <Assembly Value="OpenSpan.Runtime.Core, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" />
      <Assembly Value="OpenSpan.Script, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" />
      <Assembly Value="OSComponents.Utilities.TestHarness, Version=8.0.21.0, Culture=neutral, PublicKeyToken=c0cb69f111622a50" />
    </AssemblyReferences>
    <DynamicAssemblyReferences />
    <FileReferences />
    <BuildReferences />
  </ComponentInfo>
  <Component Version="1.0">
    <OpenSpan.Automation.GlobalContainer Name="_Demo_GC" Id="GlobalContainer-8D53AA1A011B17D" />
    <OpenSpan.Office.MicrosoftExcel Name="microsoftExcel1" Id="MicrosoftExcel-8D53AA1A49B9D1B">
      <Mode Value="Open" />
      <StartOnProjectStart Value="True" />
      <SupportedVersion Value="14" />
      <Workbook Value="C:\Users\bolltaxa\Documents\Pega Robotics Studio\Projects\pega-pras-excel-demo\Agents_By_Node_Type.xlsx" />
    </OpenSpan.Office.MicrosoftExcel>
    <OSComponents.Utilities.TestHarness.TestHarness Name="testHarness1" Id="TestHarness-8D53AEC674CC310">
      <AutomationHistoryCount Value="10" />
      <Exceptions Value="True" />
      <KeepOpen Value="True" />
      <Logging Value="False" />
      <TopMost Value="True" />
      <WinHllapiDllName Value="" />
    </OSComponents.Utilities.TestHarness.TestHarness>
    <OpenSpan.Controls.StringUtils Name="stringUtils1" Id="StringUtils-8D53B0768732E30" />
    <OpenSpan.Controls.MessageDialog Name="messageDialog1" Id="MessageDialog-8D540B2BCAA934A">
      <Caption Value="Information" />
    </OpenSpan.Controls.MessageDialog>
    <OpenSpan.Script.Custom.Script Name="HelperScripts" Id="Script-8D5B65FEC42AD8F">
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicMethodInfo dynamicType="Method" name="NumberToCell" aliasName="NumberToCell" visibility="DefaultOn" source="" blockTypeName="" returnType="System.String">
            <param name="number" aliasName="number" paramType="System.Int32" isIn="False" isOut="False" position="0" />
          </OpenSpan.DynamicMembers.DynamicMethodInfo>
        </Items>
      </Content>
    </OpenSpan.Script.Custom.Script>
  </Component>
</OpenSpanDesignDocument>