<OpenSpanDesignDocument Version="8.0.1026.0" Serializer="2.0" Culture="en-US">
  <ComponentInfo>
    <Type Value="OpenSpan.Automation.Automator" />
    <Assembly Value="OpenSpan.Automation" />
    <AssemblyReferences>
      <Assembly Value="mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" />
      <Assembly Value="System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" />
      <Assembly Value="System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" />
      <Assembly Value="Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" />
      <Assembly Value="OpenSpan, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" />
      <Assembly Value="OpenSpan.Automation, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" />
      <Assembly Value="OpenSpan.Controls, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" />
      <Assembly Value="OpenSpan.Office, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" />
    </AssemblyReferences>
    <DynamicAssemblyReferences />
    <FileReferences />
    <BuildReferences />
  </ComponentInfo>
  <Component Version="1.0">
    <OpenSpan.Automation.Automator Name="Demo_P_CopyRangeFromSheet" Id="Automator-8D53AF13F44C275">
      <AutomationDocument>
        <Name Value="" />
        <Size Value="5100, 5100" />
        <Objects>
          <ConnectionBlock>
            <ConnectableUniqueId Value="Automator-8D53AF13F44C275\EntryPoint-8D53AF1418A0666" />
            <Left Value="23" />
            <Top Value="62" />
            <PartID Value="1" />
          </ConnectionBlock>
          <ConnectionBlock type="OpenSpan.Automation.Design.ConnectionBlocks.EntryPointExecuteBlock">
            <DisplayName Value="Execute" />
            <ConnectableUniqueId Value="Automator-8D53AF13F44C275\ConnectableMethod-8D53AF58D567020" />
            <PartID Value="9" />
            <Left Value="560" />
            <Top Value="60" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="Demo_P_Cast_To_WorkSheet" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Proxy" />
            <ConnectableUniqueId Value="Automator-8D53AF13F44C275\ConnectableTypeProxy-8D53AF5920A0D20" />
            <PartID Value="11" />
            <Left Value="560" />
            <Top Value="180" />
            <Collapsed Value="False" />
            <WillExecute Value="False" />
            <InstanceName Value="_WorksheetProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D53AF13F44C275\ConnectableProperties-8D53AF5A0DA5170" />
            <PartID Value="13" />
            <Left Value="840" />
            <Top Value="60" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="_WorksheetProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Proxy" />
            <ConnectableUniqueId Value="Automator-8D53AF13F44C275\ConnectableTypeProxy-8D53AF5C85BB0E0" />
            <PartID Value="19" />
            <Left Value="840" />
            <Top Value="140" />
            <Collapsed Value="False" />
            <WillExecute Value="False" />
            <InstanceName Value="range" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D53AF13F44C275\ConnectableProperties-8D53B074DF876A0" />
            <PartID Value="21" />
            <Left Value="1080" />
            <Top Value="60" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="range" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Show" />
            <ConnectableUniqueId Value="Automator-8D53AF13F44C275\ConnectableMethod-8D53B075CCA1A80" />
            <PartID Value="22" />
            <Left Value="1560" />
            <Top Value="60" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="messageDialog1" />
            <Fittings>
              <Result Collapsed="False" ActualText="Result" />
            </Fittings>
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Proxy" />
            <ConnectableUniqueId Value="Automator-8D53AF13F44C275\ConnectableTypeProxy-8D53B099663F980" />
            <PartID Value="30" />
            <Left Value="1220" />
            <Top Value="100" />
            <Collapsed Value="False" />
            <WillExecute Value="False" />
            <InstanceName Value="colRange" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Proxy" />
            <ConnectableUniqueId Value="Automator-8D53AF13F44C275\ConnectableTypeProxy-8D53B0997A468C0" />
            <PartID Value="32" />
            <Left Value="1220" />
            <Top Value="180" />
            <Collapsed Value="False" />
            <WillExecute Value="False" />
            <InstanceName Value="rowRange" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D53AF13F44C275\ConnectableProperties-8D53B09BFC340B0" />
            <PartID Value="34" />
            <Left Value="1360" />
            <Top Value="140" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="colRange" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D53AF13F44C275\ConnectableProperties-8D53B09CF735DB0" />
            <PartID Value="36" />
            <Left Value="1400" />
            <Top Value="220" />
            <Collapsed Value="False" />
            <WillExecute Value="False" />
            <InstanceName Value="rowRange" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D53AF13F44C275\ConnectableProperties-8D53B14AB178030" />
            <PartID Value="50" />
            <Left Value="180" />
            <Top Value="60" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="microsoftExcel1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Proxy" />
            <ConnectableUniqueId Value="Automator-8D53AF13F44C275\ConnectableTypeProxy-8D53B14AEFF8FD0" />
            <PartID Value="52" />
            <Left Value="180" />
            <Top Value="140" />
            <Collapsed Value="False" />
            <WillExecute Value="False" />
            <InstanceName Value="_WorkbookProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D53AF13F44C275\ConnectableProperties-8D53B14D0664830" />
            <PartID Value="54" />
            <Left Value="360" />
            <Top Value="60" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="_WorkbookProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Proxy" />
            <ConnectableUniqueId Value="Automator-8D53AF13F44C275\ConnectableTypeProxy-8D53B14E8216A90" />
            <PartID Value="57" />
            <Left Value="360" />
            <Top Value="140" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="objectProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
        </Objects>
        <Links>
          <Link PartID="12" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="9" PortName="Result" PortType="Property" ConnectableId="Automator-8D53AF13F44C275\ConnectableMethod-8D53AF58D567020" MemberComponentId="Automator-8D53AF13F44C275\ConnectableMethod-8D53AF58D567020" />
            <To PartID="11" PortName="Instance" PortType="Property" ConnectableId="Automator-8D53AF13F44C275\ConnectableTypeProxy-8D53AF5920A0D20" MemberComponentId="Automator-8D53AF13F44C275\TypeProxy-8D53AF59205A050" />
            <LinkPoints>
              <Point value="797, 140" />
              <Point value="807, 140" />
              <Point value="812, 140" />
              <Point value="812, 156" />
              <Point value="556, 156" />
              <Point value="556, 225" />
              <Point value="555, 225" />
              <Point value="565, 225" />
            </LinkPoints>
          </Link>
          <Link PartID="15" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="9" PortName="Complete" PortType="Event" ConnectableId="Automator-8D53AF13F44C275\ConnectableMethod-8D53AF58D567020" MemberComponentId="Automator-8D53AA3CD4DFDD0\ExitPoint-8D53AA3DB396060" />
            <To PartID="13" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AF13F44C275\ConnectableProperties-8D53AF5A0DA5170" MemberComponentId="Automator-8D53AF13F44C275\ConnectableProperties-8D53AF5A0DA5170" />
            <LinkPoints>
              <Point value="797, 106" />
              <Point value="807, 106" />
              <Point value="812, 106" />
              <Point value="812, 89" />
              <Point value="835, 89" />
              <Point value="845, 89" />
            </LinkPoints>
          </Link>
          <Link PartID="20" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="13" PortName="UsedRange" PortType="Property" ConnectableId="Automator-8D53AF13F44C275\ConnectableProperties-8D53AF5A0DA5170" MemberComponentId="Automator-8D53AF13F44C275\TypeProxy-8D53AF59205A050" />
            <To PartID="19" PortName="Instance" PortType="Property" ConnectableId="Automator-8D53AF13F44C275\ConnectableTypeProxy-8D53AF5C85BB0E0" MemberComponentId="Automator-8D53AF13F44C275\TypeProxy-8D53AF5C8560B90" />
            <LinkPoints>
              <Point value="1005, 106" />
              <Point value="1015, 106" />
              <Point value="1020, 106" />
              <Point value="1020, 124" />
              <Point value="836, 124" />
              <Point value="836, 185" />
              <Point value="835, 185" />
              <Point value="845, 185" />
            </LinkPoints>
          </Link>
          <Link PartID="31" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="21" PortName="Columns" PortType="Property" ConnectableId="Automator-8D53AF13F44C275\ConnectableProperties-8D53B074DF876A0" MemberComponentId="Automator-8D53AF13F44C275\TypeProxy-8D53AF5C8560B90" />
            <To PartID="30" PortName="Instance" PortType="Property" ConnectableId="Automator-8D53AF13F44C275\ConnectableTypeProxy-8D53B099663F980" MemberComponentId="Automator-8D53AF13F44C275\TypeProxy-8D53B09966028F0" />
            <LinkPoints>
              <Point value="1189, 106" />
              <Point value="1199, 106" />
              <Point value="1204, 106" />
              <Point value="1204, 145" />
              <Point value="1215, 145" />
              <Point value="1225, 145" />
            </LinkPoints>
          </Link>
          <Link PartID="33" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="21" PortName="Rows" PortType="Property" ConnectableId="Automator-8D53AF13F44C275\ConnectableProperties-8D53B074DF876A0" MemberComponentId="Automator-8D53AF13F44C275\TypeProxy-8D53AF5C8560B90" />
            <To PartID="32" PortName="Instance" PortType="Property" ConnectableId="Automator-8D53AF13F44C275\ConnectableTypeProxy-8D53B0997A468C0" MemberComponentId="Automator-8D53AF13F44C275\TypeProxy-8D53B0997A09830" />
            <LinkPoints>
              <Point value="1189, 123" />
              <Point value="1199, 123" />
              <Point value="1204, 123" />
              <Point value="1204, 225" />
              <Point value="1215, 225" />
              <Point value="1225, 225" />
            </LinkPoints>
          </Link>
          <Link PartID="42" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="34" PortName="Count" PortType="Property" ConnectableId="Automator-8D53AF13F44C275\ConnectableProperties-8D53B09BFC340B0" MemberComponentId="Automator-8D53AF13F44C275\TypeProxy-8D53B09966028F0" />
            <To PartID="22" PortName="message" PortType="Property" ConnectableId="Automator-8D53AF13F44C275\ConnectableMethod-8D53B075CCA1A80" MemberComponentId="Automator-8D53AF13F44C275\ConnectableMethod-8D53B075CCA1A80" />
            <LinkPoints>
              <Point value="1469, 186" />
              <Point value="1479, 186" />
              <Point value="1484, 186" />
              <Point value="1484, 106" />
              <Point value="1555, 106" />
              <Point value="1565, 106" />
            </LinkPoints>
          </Link>
          <Link PartID="43" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="21" PortName="Complete" PortType="Event" ConnectableId="Automator-8D53AF13F44C275\ConnectableProperties-8D53B074DF876A0" MemberComponentId="Automator-8D53AF13F44C275\ConnectableProperties-8D53B074DF876A0" />
            <To PartID="22" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AF13F44C275\ConnectableMethod-8D53B075CCA1A80" MemberComponentId="Automator-8D53AF13F44C275\ConnectableMethod-8D53B075CCA1A80" />
            <LinkPoints>
              <Point value="1189, 89" />
              <Point value="1199, 89" />
              <Point value="1199, 89" />
              <Point value="1199, 89" />
              <Point value="1555, 89" />
              <Point value="1565, 89" />
            </LinkPoints>
          </Link>
          <Link PartID="44" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="13" PortName="Complete" PortType="Event" ConnectableId="Automator-8D53AF13F44C275\ConnectableProperties-8D53AF5A0DA5170" MemberComponentId="Automator-8D53AF13F44C275\ConnectableProperties-8D53AF5A0DA5170" />
            <To PartID="21" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AF13F44C275\ConnectableProperties-8D53B074DF876A0" MemberComponentId="Automator-8D53AF13F44C275\ConnectableProperties-8D53B074DF876A0" />
            <LinkPoints>
              <Point value="1005, 89" />
              <Point value="1015, 89" />
              <Point value="1045, 89" />
              <Point value="1045, 89" />
              <Point value="1075, 89" />
              <Point value="1085, 89" />
            </LinkPoints>
          </Link>
          <Link PartID="51" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="1" PortName="Complete" PortType="Event" ConnectableId="Automator-8D53AF13F44C275\EntryPoint-8D53AF1418A0666" MemberComponentId="Automator-8D53AF13F44C275\EntryPoint-8D53AF1418A0666" />
            <To PartID="50" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AF13F44C275\ConnectableProperties-8D53B14AB178030" MemberComponentId="Automator-8D53AF13F44C275\ConnectableProperties-8D53B14AB178030" />
            <LinkPoints>
              <Point value="134, 80" />
              <Point value="144, 80" />
              <Point value="148, 80" />
              <Point value="148, 89" />
              <Point value="175, 89" />
              <Point value="185, 89" />
            </LinkPoints>
          </Link>
          <Link PartID="53" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="50" PortName="ExcelWorkbook" PortType="Property" ConnectableId="Automator-8D53AF13F44C275\ConnectableProperties-8D53B14AB178030" MemberComponentId="GlobalContainer-8D53AA1A011B17D\MicrosoftExcel-8D53AA1A49B9D1B" />
            <To PartID="52" PortName="Instance" PortType="Property" ConnectableId="Automator-8D53AF13F44C275\ConnectableTypeProxy-8D53B14AEFF8FD0" MemberComponentId="Automator-8D53AF13F44C275\TypeProxy-8D53B14AEFBE650" />
            <LinkPoints>
              <Point value="326, 106" />
              <Point value="336, 106" />
              <Point value="340, 106" />
              <Point value="340, 124" />
              <Point value="172, 124" />
              <Point value="172, 185" />
              <Point value="175, 185" />
              <Point value="185, 185" />
            </LinkPoints>
          </Link>
          <Link PartID="55" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="50" PortName="Complete" PortType="Event" ConnectableId="Automator-8D53AF13F44C275\ConnectableProperties-8D53B14AB178030" MemberComponentId="Automator-8D53AF13F44C275\ConnectableProperties-8D53B14AB178030" />
            <To PartID="54" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AF13F44C275\ConnectableProperties-8D53B14D0664830" MemberComponentId="Automator-8D53AF13F44C275\ConnectableProperties-8D53B14D0664830" />
            <LinkPoints>
              <Point value="326, 89" />
              <Point value="336, 89" />
              <Point value="346, 89" />
              <Point value="346, 89" />
              <Point value="355, 89" />
              <Point value="365, 89" />
            </LinkPoints>
          </Link>
          <Link PartID="56" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="54" PortName="Complete" PortType="Event" ConnectableId="Automator-8D53AF13F44C275\ConnectableProperties-8D53B14D0664830" MemberComponentId="Automator-8D53AF13F44C275\ConnectableProperties-8D53B14D0664830" />
            <To PartID="9" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AF13F44C275\ConnectableMethod-8D53AF58D567020" MemberComponentId="Automator-8D53AF13F44C275\ConnectableMethod-8D53AF58D567020" />
            <LinkPoints>
              <Point value="521, 89" />
              <Point value="531, 89" />
              <Point value="543, 89" />
              <Point value="543, 89" />
              <Point value="555, 89" />
              <Point value="565, 89" />
            </LinkPoints>
          </Link>
          <Link PartID="58" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="54" PortName="ActiveSheet" PortType="Property" ConnectableId="Automator-8D53AF13F44C275\ConnectableProperties-8D53B14D0664830" MemberComponentId="Automator-8D53AF13F44C275\TypeProxy-8D53B14AEFBE650" />
            <To PartID="57" PortName="Instance" PortType="Property" ConnectableId="Automator-8D53AF13F44C275\ConnectableTypeProxy-8D53B14E8216A90" MemberComponentId="Automator-8D53AF13F44C275\TypeProxy-8D53B14E81E0F30" />
            <LinkPoints>
              <Point value="521, 106" />
              <Point value="531, 106" />
              <Point value="532, 106" />
              <Point value="532, 124" />
              <Point value="356, 124" />
              <Point value="356, 185" />
              <Point value="355, 185" />
              <Point value="365, 185" />
            </LinkPoints>
          </Link>
          <Link PartID="59" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="57" PortName="Instance" PortType="Property" ConnectableId="Automator-8D53AF13F44C275\ConnectableTypeProxy-8D53B14E8216A90" MemberComponentId="Automator-8D53AF13F44C275\TypeProxy-8D53B14E81E0F30" />
            <To PartID="9" PortName="param1" PortType="Property" ConnectableId="Automator-8D53AF13F44C275\ConnectableMethod-8D53AF58D567020" MemberComponentId="Automator-8D53AF13F44C275\ConnectableMethod-8D53AF58D567020" />
            <LinkPoints>
              <Point value="488, 185" />
              <Point value="498, 185" />
              <Point value="500, 185" />
              <Point value="500, 123" />
              <Point value="555, 123" />
              <Point value="565, 123" />
            </LinkPoints>
          </Link>
        </Links>
        <Comments />
        <SubGraphs />
      </AutomationDocument>
      <DocumentPosition Value="Binary">
        <Binary>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj00LjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5Qb2ludEYCAAAAAXgBeQAACwsCAAAAAACMQwAAgEEL</Binary>
      </DocumentPosition>
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicMethodInfo dynamicType="Method" name="_EntryPointExecute" aliasName="Execute" visibility="DefaultOn" source="" blockTypeName="OpenSpan.Automation.Design.ConnectionBlocks.EntryPointExecuteBlock" returnType="System.Void" />
        </Items>
      </Content>
    </OpenSpan.Automation.Automator>
    <OpenSpan.Automation.EntryPoint Name="entryPoint1" Id="EntryPoint-8D53AF1418A0666">
      <AliasName Value="Execute" />
      <ComponentName Value="Execute" />
      <DisplayName Value="Execute" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Automation.EntryPoint" />
      <InstanceUniqueId Value="Automator-8D53AF13F44C275\EntryPoint-8D53AF1418A0666" />
      <MemberDetails Value="" />
      <MethodName Value="_EntryPointExecute" />
      <Removing Value="False" />
      <UseAlias Value="True" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Result" canRead="True" canWrite="True" type="System.Void" aliasName="Result" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Automation.EntryPoint>
    <OpenSpan.Automation.ConnectableMethod Name="connectableMethod1" Id="ConnectableMethod-8D53AF58D567020">
      <ComponentName Value="Demo_P_Cast_To_WorkSheet" />
      <DisplayName Value="Execute" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Automation.Automator" />
      <InstanceUniqueId Value="Automator-8D53AA3CD4DFDD0" />
      <MemberDetails Value=".Execute() Method" />
      <ParamsLength Value="0" />
      <SerializedParamsDefaultValues Value="" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Result" canRead="True" canWrite="False" typeAssembly="Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" type="Microsoft.Office.Interop.Excel._Worksheet" aliasName="Result" shouldSerialize="False" visibility="DefaultOn" source="" blockTypeName="" />
        </Items>
      </Content>
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="_EntryPointExecute" />
            <MemberType Value="Method" />
            <TypeAssemblyName Value="Microsoft.Office.Interop.Excel" />
            <TypeName Value="Microsoft.Office.Interop.Excel._Worksheet" />
            <Content Name="Signature">
              <OpenSpan.Automation.MethodSignature>
                <ReturnType Value="Microsoft.Office.Interop.Excel._Worksheet" />
                <Content Name="ParameterPrototype">
                  <Items>
                    <OpenSpan.Automation.ParameterPrototype>
                      <CanRead Value="False" />
                      <CanWrite Value="True" />
                      <DefaultSet Value="False" />
                      <ParamName Value="param1" />
                      <Position Value="0" />
                      <TypeAssemblyName Value="Microsoft.Office.Interop.Excel" />
                      <TypeName Value="Microsoft.Office.Interop.Excel._Worksheet" />
                    </OpenSpan.Automation.ParameterPrototype>
                  </Items>
                </Content>
              </OpenSpan.Automation.MethodSignature>
            </Content>
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableMethod>
    <OpenSpan.Design.TypeProxy Name="_WorksheetProxy1" Id="TypeProxy-8D53AF59205A050">
      <AliasName Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel._Worksheet, Microsoft.Office.Interop.Excel" />
      <UseAlias Value="False" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Instance" canRead="True" canWrite="True" typeAssembly="Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" type="Microsoft.Office.Interop.Excel._Worksheet" aliasName="Instance" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Design.TypeProxy>
    <OpenSpan.Automation.ConnectableTypeProxy Name="connectableTypeProxy3" Id="ConnectableTypeProxy-8D53AF5920A0D20">
      <ComponentName Value="_WorksheetProxy1" />
      <DisplayName Value="Proxy" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel._Worksheet" />
      <InstanceUniqueId Value="Automator-8D53AF13F44C275\TypeProxy-8D53AF59205A050" />
      <MemberDetails Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel._Worksheet" />
    </OpenSpan.Automation.ConnectableTypeProxy>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties3" Id="ConnectableProperties-8D53AF5A0DA5170">
      <ComponentName Value="_WorksheetProxy1" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel._Worksheet" />
      <InstanceUniqueId Value="Automator-8D53AF13F44C275\TypeProxy-8D53AF59205A050" />
      <MemberDetails Value=".UsedRange Property" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="UsedRange" />
            <MemberType Value="Property" />
            <TypeAssemblyName Value="Microsoft.Office.Interop.Excel" />
            <TypeName Value="Microsoft.Office.Interop.Excel.Range" />
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableProperties>
    <OpenSpan.Design.TypeProxy Name="range" Id="TypeProxy-8D53AF5C8560B90">
      <AliasName Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel.Range, Microsoft.Office.Interop.Excel" />
      <UseAlias Value="False" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Instance" canRead="True" canWrite="True" typeAssembly="Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" type="Microsoft.Office.Interop.Excel.Range" aliasName="Instance" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Design.TypeProxy>
    <OpenSpan.Automation.ConnectableTypeProxy Name="connectableTypeProxy4" Id="ConnectableTypeProxy-8D53AF5C85BB0E0">
      <ComponentName Value="rangeProxy1" />
      <DisplayName Value="Proxy" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel.Range" />
      <InstanceUniqueId Value="Automator-8D53AF13F44C275\TypeProxy-8D53AF5C8560B90" />
      <MemberDetails Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel.Range" />
    </OpenSpan.Automation.ConnectableTypeProxy>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties4" Id="ConnectableProperties-8D53B074DF876A0">
      <ComponentName Value="rangeProxy1" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel.Range" />
      <InstanceUniqueId Value="Automator-8D53AF13F44C275\TypeProxy-8D53AF5C8560B90" />
      <MemberDetails Value=" - Properties(Columns, Rows)" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="Columns" />
            <MemberType Value="Property" />
            <TypeAssemblyName Value="Microsoft.Office.Interop.Excel" />
            <TypeName Value="Microsoft.Office.Interop.Excel.Range" />
          </OpenSpan.Automation.MemberPrototype>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="Rows" />
            <MemberType Value="Property" />
            <TypeAssemblyName Value="Microsoft.Office.Interop.Excel" />
            <TypeName Value="Microsoft.Office.Interop.Excel.Range" />
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableProperties>
    <OpenSpan.Controls.MessageDialog Name="messageDialog1" Id="MessageDialog-8D53B075BB66090">
      <Caption Value="Information" />
      <Scope Value="Local" Extended="True" />
    </OpenSpan.Controls.MessageDialog>
    <OpenSpan.Automation.ConnectableMethod Name="connectableMethod2" Id="ConnectableMethod-8D53B075CCA1A80">
      <ComponentName Value="messageDialog1" />
      <DisplayName Value="Show" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Controls.MessageDialog" />
      <InstanceUniqueId Value="Automator-8D53AF13F44C275\MessageDialog-8D53B075BB66090" />
      <MemberDetails Value=".Show() Method" />
      <ParamsLength Value="0" />
      <SerializedParamsDefaultValues Value="" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Result" canRead="True" canWrite="False" type="System.Windows.Forms.DialogResult" aliasName="Result" shouldSerialize="False" visibility="DefaultOn" source="" blockTypeName="" />
        </Items>
      </Content>
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="Show" />
            <MemberType Value="Method" />
            <TypeAssemblyName Value="System.Windows.Forms" />
            <TypeName Value="System.Windows.Forms.DialogResult" />
            <Content Name="Signature">
              <OpenSpan.Automation.MethodSignature>
                <ReturnType Value="System.Windows.Forms.DialogResult" />
                <Content Name="ParameterPrototype">
                  <Items>
                    <OpenSpan.Automation.ParameterPrototype>
                      <CanRead Value="False" />
                      <CanWrite Value="True" />
                      <DefaultSet Value="False" />
                      <ParamName Value="message" />
                      <Position Value="0" />
                      <TypeName Value="System.String" />
                    </OpenSpan.Automation.ParameterPrototype>
                  </Items>
                </Content>
              </OpenSpan.Automation.MethodSignature>
            </Content>
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableMethod>
    <OpenSpan.Design.TypeProxy Name="colRange" Id="TypeProxy-8D53B09966028F0">
      <AliasName Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel.Range, Microsoft.Office.Interop.Excel" />
      <UseAlias Value="False" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Instance" canRead="True" canWrite="True" typeAssembly="Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" type="Microsoft.Office.Interop.Excel.Range" aliasName="Instance" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Design.TypeProxy>
    <OpenSpan.Automation.ConnectableTypeProxy Name="connectableTypeProxy5" Id="ConnectableTypeProxy-8D53B099663F980">
      <ComponentName Value="rangeProxy2" />
      <DisplayName Value="Proxy" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Design.TypeProxy" />
      <InstanceUniqueId Value="Automator-8D53AF13F44C275\TypeProxy-8D53B09966028F0" />
      <MemberDetails Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel.Range" />
    </OpenSpan.Automation.ConnectableTypeProxy>
    <OpenSpan.Design.TypeProxy Name="rowRange" Id="TypeProxy-8D53B0997A09830">
      <AliasName Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel.Range, Microsoft.Office.Interop.Excel" />
      <UseAlias Value="False" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Instance" canRead="True" canWrite="True" typeAssembly="Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" type="Microsoft.Office.Interop.Excel.Range" aliasName="Instance" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Design.TypeProxy>
    <OpenSpan.Automation.ConnectableTypeProxy Name="connectableTypeProxy6" Id="ConnectableTypeProxy-8D53B0997A468C0">
      <ComponentName Value="rangeProxy3" />
      <DisplayName Value="Proxy" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Design.TypeProxy" />
      <InstanceUniqueId Value="Automator-8D53AF13F44C275\TypeProxy-8D53B0997A09830" />
      <MemberDetails Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel.Range" />
    </OpenSpan.Automation.ConnectableTypeProxy>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties5" Id="ConnectableProperties-8D53B09BFC340B0">
      <ComponentName Value="colRange" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel.Range" />
      <InstanceUniqueId Value="Automator-8D53AF13F44C275\TypeProxy-8D53B09966028F0" />
      <MemberDetails Value=".Count Property" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="Count" />
            <MemberType Value="Property" />
            <TypeName Value="System.Int32" />
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableProperties>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties6" Id="ConnectableProperties-8D53B09CF735DB0">
      <ComponentName Value="rowRange" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel.Range" />
      <InstanceUniqueId Value="Automator-8D53AF13F44C275\TypeProxy-8D53B0997A09830" />
      <MemberDetails Value=".Count Property" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="Count" />
            <MemberType Value="Property" />
            <TypeName Value="System.Int32" />
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableProperties>
    <OpenSpan.Design.TypeProxy Name="expandableFilePathProxy1" Id="TypeProxy-8D53B14874CE730">
      <AliasName Value="" />
      <ProxiedTypeName Value="OpenSpan.ComponentModel.ExpandableFilePath, OpenSpan" />
      <UseAlias Value="False" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Instance" canRead="True" canWrite="True" typeAssembly="OpenSpan, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" type="OpenSpan.ComponentModel.ExpandableFilePath" aliasName="Instance" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Design.TypeProxy>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties1" Id="ConnectableProperties-8D53B14AB178030">
      <ComponentName Value="microsoftExcel1" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Office.MicrosoftExcel" />
      <InstanceUniqueId Value="GlobalContainer-8D53AA1A011B17D\MicrosoftExcel-8D53AA1A49B9D1B" />
      <MemberDetails Value=".ExcelWorkbook Property" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="ExcelWorkbook" />
            <MemberType Value="Property" />
            <TypeAssemblyName Value="Microsoft.Office.Interop.Excel" />
            <TypeName Value="Microsoft.Office.Interop.Excel._Workbook" />
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableProperties>
    <OpenSpan.Design.TypeProxy Name="_WorkbookProxy1" Id="TypeProxy-8D53B14AEFBE650">
      <AliasName Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel._Workbook, Microsoft.Office.Interop.Excel" />
      <UseAlias Value="False" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Instance" canRead="True" canWrite="True" typeAssembly="Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" type="Microsoft.Office.Interop.Excel._Workbook" aliasName="Instance" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Design.TypeProxy>
    <OpenSpan.Automation.ConnectableTypeProxy Name="connectableTypeProxy1" Id="ConnectableTypeProxy-8D53B14AEFF8FD0">
      <ComponentName Value="_WorkbookProxy1" />
      <DisplayName Value="Proxy" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Design.TypeProxy" />
      <InstanceUniqueId Value="Automator-8D53AF13F44C275\TypeProxy-8D53B14AEFBE650" />
      <MemberDetails Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel._Workbook" />
    </OpenSpan.Automation.ConnectableTypeProxy>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties2" Id="ConnectableProperties-8D53B14D0664830">
      <ComponentName Value="_WorkbookProxy1" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel._Workbook" />
      <InstanceUniqueId Value="Automator-8D53AF13F44C275\TypeProxy-8D53B14AEFBE650" />
      <MemberDetails Value=".ActiveSheet Property" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="ActiveSheet" />
            <MemberType Value="Property" />
            <TypeName Value="System.Object" />
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableProperties>
    <OpenSpan.Design.TypeProxy Name="objectProxy1" Id="TypeProxy-8D53B14E81E0F30">
      <AliasName Value="" />
      <ProxiedTypeName Value="System.Object, mscorlib" />
      <UseAlias Value="False" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Instance" canRead="True" canWrite="True" type="System.Object" aliasName="Instance" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Design.TypeProxy>
    <OpenSpan.Automation.ConnectableTypeProxy Name="connectableTypeProxy2" Id="ConnectableTypeProxy-8D53B14E8216A90">
      <ComponentName Value="objectProxy1" />
      <DisplayName Value="Proxy" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Design.TypeProxy" />
      <InstanceUniqueId Value="Automator-8D53AF13F44C275\TypeProxy-8D53B14E81E0F30" />
      <MemberDetails Value="" />
      <ProxiedTypeName Value="System.Object" />
    </OpenSpan.Automation.ConnectableTypeProxy>
  </Component>
</OpenSpanDesignDocument>