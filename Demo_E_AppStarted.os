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
    <OpenSpan.Automation.Automator Name="Demo_E_AppStarted" Id="Automator-8D53AA217ABD62C">
      <AutomationDocument>
        <Name Value="" />
        <Size Value="5000, 5000" />
        <Objects>
          <ConnectionBlock>
            <DisplayName Value="DesignForm.Started" />
            <ConnectableUniqueId Value="Automator-8D53AA217ABD62C\ConnectableEvent-8D53AA222A62D1C" />
            <PartID Value="1" />
            <Left Value="20" />
            <Top Value="40" />
            <Collapsed Value="True" />
            <WillExecute Value="True" />
            <InstanceName Value="ExcelQuery" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D53AA217ABD62C\ConnectableProperties-8D53AA293B5FE20" />
            <PartID Value="6" />
            <Left Value="200" />
            <Top Value="40" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="microsoftExcel1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Proxy" />
            <ConnectableUniqueId Value="Automator-8D53AA217ABD62C\ConnectableTypeProxy-8D53AA2984CC450" />
            <PartID Value="8" />
            <Left Value="388" />
            <Top Value="91" />
            <Collapsed Value="False" />
            <WillExecute Value="False" />
            <InstanceName Value="_WorkbookProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D53AA217ABD62C\ConnectableProperties-8D53AA33ACF3EB0" />
            <PartID Value="12" />
            <Left Value="180" />
            <Top Value="200" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="_WorkbookProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="StartLoop" />
            <ConnectableUniqueId Value="Automator-8D53AA217ABD62C\ListLoop-8D53AA34B20A4C0" />
            <PartID Value="13" />
            <Left Value="400" />
            <Top Value="200" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="listLoop1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Clear" />
            <ConnectableUniqueId Value="Automator-8D53AA217ABD62C\ConnectableMethod-8D53AA3780B3590" />
            <PartID Value="17" />
            <Left Value="20" />
            <Top Value="120" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="cmbTabs.Items" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Add" />
            <ConnectableUniqueId Value="Automator-8D53AA217ABD62C\ConnectableMethod-8D53AA37A610C70" />
            <PartID Value="18" />
            <Left Value="820" />
            <Top Value="320" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="cmbTabs.Items" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock type="OpenSpan.Automation.Design.ConnectionBlocks.EntryPointExecuteBlock">
            <DisplayName Value="Execute" />
            <ConnectableUniqueId Value="Automator-8D53AA217ABD62C\ConnectableMethod-8D53AA3F4886930" />
            <PartID Value="27" />
            <Left Value="580" />
            <Top Value="200" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="Demo_P_Cast_To_WorkSheet" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Proxy" />
            <ConnectableUniqueId Value="Automator-8D53AA217ABD62C\ConnectableTypeProxy-8D53AA40FD3B730" />
            <PartID Value="32" />
            <Left Value="580" />
            <Top Value="320" />
            <Collapsed Value="False" />
            <WillExecute Value="False" />
            <InstanceName Value="_WorksheetProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D53AA217ABD62C\ConnectableProperties-8D53AA41B347740" />
            <PartID Value="34" />
            <Left Value="580" />
            <Top Value="400" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="_WorksheetProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
        </Objects>
        <Links>
          <Link PartID="9" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="6" PortName="ExcelWorkbook" PortType="Property" ConnectableId="Automator-8D53AA217ABD62C\ConnectableProperties-8D53AA293B5FE20" MemberComponentId="GlobalContainer-8D53AA1A011B17D\MicrosoftExcel-8D53AA1A49B9D1B" />
            <To PartID="8" PortName="Instance" PortType="Property" ConnectableId="Automator-8D53AA217ABD62C\ConnectableTypeProxy-8D53AA2984CC450" MemberComponentId="Automator-8D53AA217ABD62C\TypeProxy-8D53AA298483070" />
            <LinkPoints>
              <Point value="346, 86" />
              <Point value="356, 86" />
              <Point value="356, 86" />
              <Point value="356, 136" />
              <Point value="383, 136" />
              <Point value="393, 136" />
            </LinkPoints>
          </Link>
          <Link PartID="15" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="12" PortName="Worksheets" PortType="Property" ConnectableId="Automator-8D53AA217ABD62C\ConnectableProperties-8D53AA33ACF3EB0" MemberComponentId="Automator-8D53AA217ABD62C\TypeProxy-8D53AA298483070" />
            <To PartID="13" PortName="List" PortType="Property" ConnectableId="Automator-8D53AA217ABD62C\ListLoop-8D53AA34B20A4C0" MemberComponentId="Automator-8D53AA217ABD62C\ListLoop-8D53AA34B20A4C0" />
            <LinkPoints>
              <Point value="341, 246" />
              <Point value="351, 246" />
              <Point value="351, 246" />
              <Point value="351, 246" />
              <Point value="395, 246" />
              <Point value="405, 246" />
            </LinkPoints>
          </Link>
          <Link PartID="19" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="1" PortName="Raised" PortType="Event" ConnectableId="Automator-8D53AA217ABD62C\ConnectableEvent-8D53AA222A62D1C" MemberComponentId="Automator-8D53AA217ABD62C\ConnectableEvent-8D53AA222A62D1C" />
            <To PartID="17" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AA217ABD62C\ConnectableMethod-8D53AA3780B3590" MemberComponentId="Automator-8D53AA217ABD62C\ConnectableMethod-8D53AA3780B3590" />
            <LinkPoints>
              <Point value="140, 69" />
              <Point value="150, 69" />
              <Point value="150, 109" />
              <Point value="15, 109" />
              <Point value="15, 149" />
              <Point value="25, 149" />
            </LinkPoints>
          </Link>
          <Link PartID="20" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="17" PortName="Complete" PortType="Event" ConnectableId="Automator-8D53AA217ABD62C\ConnectableMethod-8D53AA3780B3590" MemberComponentId="Automator-8D53AA217ABD62C\ConnectableMethod-8D53AA3780B3590" />
            <To PartID="6" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AA217ABD62C\ConnectableProperties-8D53AA293B5FE20" MemberComponentId="Automator-8D53AA217ABD62C\ConnectableProperties-8D53AA293B5FE20" />
            <LinkPoints>
              <Point value="163, 149" />
              <Point value="173, 149" />
              <Point value="184, 149" />
              <Point value="184, 69" />
              <Point value="195, 69" />
              <Point value="205, 69" />
            </LinkPoints>
          </Link>
          <Link PartID="21" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="6" PortName="Complete" PortType="Event" ConnectableId="Automator-8D53AA217ABD62C\ConnectableProperties-8D53AA293B5FE20" MemberComponentId="Automator-8D53AA217ABD62C\ConnectableProperties-8D53AA293B5FE20" />
            <To PartID="13" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AA217ABD62C\ListLoop-8D53AA34B20A4C0" MemberComponentId="Automator-8D53AA217ABD62C\ListLoop-8D53AA34B20A4C0" />
            <LinkPoints>
              <Point value="346, 69" />
              <Point value="356, 69" />
              <Point value="376, 69" />
              <Point value="376, 229" />
              <Point value="395, 229" />
              <Point value="405, 229" />
            </LinkPoints>
          </Link>
          <Link PartID="29" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="13" PortName="ItemYielded" PortType="Event" ConnectableId="Automator-8D53AA217ABD62C\ListLoop-8D53AA34B20A4C0" MemberComponentId="Automator-8D53AA217ABD62C\ListLoop-8D53AA34B20A4C0" />
            <To PartID="27" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AA217ABD62C\ConnectableMethod-8D53AA3F4886930" MemberComponentId="Automator-8D53AA217ABD62C\ConnectableMethod-8D53AA3F4886930" />
            <LinkPoints>
              <Point value="519, 263" />
              <Point value="529, 263" />
              <Point value="552, 263" />
              <Point value="552, 229" />
              <Point value="575, 229" />
              <Point value="585, 229" />
            </LinkPoints>
          </Link>
          <Link PartID="31" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="13" PortName="TypedItem" PortType="Property" ConnectableId="Automator-8D53AA217ABD62C\ListLoop-8D53AA34B20A4C0" MemberComponentId="Automator-8D53AA217ABD62C\ListLoop-8D53AA34B20A4C0" />
            <To PartID="27" PortName="param1" PortType="Property" ConnectableId="Automator-8D53AA217ABD62C\ConnectableMethod-8D53AA3F4886930" MemberComponentId="Automator-8D53AA217ABD62C\ConnectableMethod-8D53AA3F4886930" />
            <LinkPoints>
              <Point value="519, 280" />
              <Point value="529, 280" />
              <Point value="552, 280" />
              <Point value="552, 263" />
              <Point value="575, 263" />
              <Point value="585, 263" />
            </LinkPoints>
          </Link>
          <Link PartID="33" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="27" PortName="Result" PortType="Property" ConnectableId="Automator-8D53AA217ABD62C\ConnectableMethod-8D53AA3F4886930" MemberComponentId="Automator-8D53AA217ABD62C\ConnectableMethod-8D53AA3F4886930" />
            <To PartID="32" PortName="Instance" PortType="Property" ConnectableId="Automator-8D53AA217ABD62C\ConnectableTypeProxy-8D53AA40FD3B730" MemberComponentId="Automator-8D53AA217ABD62C\TypeProxy-8D53AA40FCF2350" />
            <LinkPoints>
              <Point value="817, 280" />
              <Point value="827, 280" />
              <Point value="827, 323" />
              <Point value="575, 323" />
              <Point value="575, 365" />
              <Point value="585, 365" />
            </LinkPoints>
          </Link>
          <Link PartID="36" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="34" PortName="Name" PortType="Property" ConnectableId="Automator-8D53AA217ABD62C\ConnectableProperties-8D53AA41B347740" MemberComponentId="Automator-8D53AA217ABD62C\TypeProxy-8D53AA40FCF2350" />
            <To PartID="18" PortName="item" PortType="Property" ConnectableId="Automator-8D53AA217ABD62C\ConnectableMethod-8D53AA37A610C70" MemberComponentId="Automator-8D53AA217ABD62C\ConnectableMethod-8D53AA37A610C70" />
            <LinkPoints>
              <Point value="745, 446" />
              <Point value="755, 446" />
              <Point value="756, 446" />
              <Point value="756, 366" />
              <Point value="815, 366" />
              <Point value="825, 366" />
            </LinkPoints>
          </Link>
          <Link PartID="37" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="27" PortName="Complete" PortType="Event" ConnectableId="Automator-8D53AA217ABD62C\ConnectableMethod-8D53AA3F4886930" MemberComponentId="Automator-8D53AA3CD4DFDD0\ExitPoint-8D53AA3DB396060" />
            <To PartID="18" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AA217ABD62C\ConnectableMethod-8D53AA37A610C70" MemberComponentId="Automator-8D53AA217ABD62C\ConnectableMethod-8D53AA37A610C70" />
            <LinkPoints>
              <Point value="817, 246" />
              <Point value="827, 246" />
              <Point value="827, 297" />
              <Point value="815, 297" />
              <Point value="815, 349" />
              <Point value="825, 349" />
            </LinkPoints>
          </Link>
        </Links>
        <Comments />
        <SubGraphs />
      </AutomationDocument>
    </OpenSpan.Automation.Automator>
    <OpenSpan.Automation.ConnectableEvent Name="connectableEvent1" Id="ConnectableEvent-8D53AA222A62D1C">
      <ComponentName Value="ExcelQuery" />
      <DisplayName Value="DesignForm.Started" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Design.DesignForm" />
      <InstanceUniqueId Value="DesignForm-8D53AA1BDDF1678" />
      <MemberDetails Value=".Started Event" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="Started" />
            <MemberType Value="Event" />
            <TypeName Value="System.EventHandler" />
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableEvent>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties1" Id="ConnectableProperties-8D53AA293B5FE20">
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
    <OpenSpan.Design.TypeProxy Name="_WorkbookProxy1" Id="TypeProxy-8D53AA298483070">
      <AliasName Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel._Workbook, Microsoft.Office.Interop.Excel" />
      <UseAlias Value="False" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Instance" canRead="True" canWrite="True" typeAssembly="Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" type="Microsoft.Office.Interop.Excel._Workbook" aliasName="Instance" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Design.TypeProxy>
    <OpenSpan.Automation.ConnectableTypeProxy Name="connectableTypeProxy1" Id="ConnectableTypeProxy-8D53AA2984CC450">
      <ComponentName Value="_WorkbookProxy1" />
      <DisplayName Value="Proxy" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel._Workbook" />
      <InstanceUniqueId Value="Automator-8D53AA217ABD62C\TypeProxy-8D53AA298483070" />
      <MemberDetails Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel._Workbook" />
    </OpenSpan.Automation.ConnectableTypeProxy>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties2" Id="ConnectableProperties-8D53AA33ACF3EB0">
      <ComponentName Value="_WorkbookProxy1" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel._Workbook" />
      <InstanceUniqueId Value="Automator-8D53AA217ABD62C\TypeProxy-8D53AA298483070" />
      <MemberDetails Value=".Worksheets Property" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="Worksheets" />
            <MemberType Value="Property" />
            <TypeAssemblyName Value="Microsoft.Office.Interop.Excel" />
            <TypeName Value="Microsoft.Office.Interop.Excel.Sheets" />
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableProperties>
    <OpenSpan.Controls.ListLoop Name="listLoop1" Id="ListLoop-8D53AA34B20A4C0">
      <ComponentName Value="listLoop1" />
      <DisplayName Value="StartLoop" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Controls.ListLoop" />
      <InstanceUniqueId Value="Automator-8D53AA217ABD62C\ListLoop-8D53AA34B20A4C0" />
      <MemberDetails Value="" />
      <Scope Value="Local" Extended="True" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="TypedItem" canRead="True" canWrite="False" type="System.Object" aliasName="TypedItem" shouldSerialize="False" visibility="DefaultOn" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Controls.ListLoop>
    <OpenSpan.Automation.ConnectableMethod Name="connectableMethod1" Id="ConnectableMethod-8D53AA3780B3590">
      <ComponentName Value="cmbTabs" />
      <DisplayName Value="Clear" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="System.Windows.Forms.ComboBox+ObjectCollection" />
      <InstanceUniqueId Value="DesignForm-8D53AA1BDDF1678\ComboBox-8D53AA1EEEFDA5E" />
      <MemberDetails Value=".Clear() Method" />
      <ParamsLength Value="0" />
      <SerializedParamsDefaultValues Value="" />
      <SubProperty Value="Items" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="Clear" />
            <MemberType Value="Method" />
            <TypeName Value="System.Void" />
            <Content Name="Signature">
              <OpenSpan.Automation.MethodSignature>
                <ReturnType Value="System.Void" />
              </OpenSpan.Automation.MethodSignature>
            </Content>
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableMethod>
    <OpenSpan.Automation.ConnectableMethod Name="connectableMethod2" Id="ConnectableMethod-8D53AA37A610C70">
      <ComponentName Value="cmbTabs" />
      <DisplayName Value="Add" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="System.Windows.Forms.ComboBox+ObjectCollection" />
      <InstanceUniqueId Value="DesignForm-8D53AA1BDDF1678\ComboBox-8D53AA1EEEFDA5E" />
      <MemberDetails Value=".Add() Method" />
      <ParamsLength Value="0" />
      <SerializedParamsDefaultValues Value="" />
      <SubProperty Value="Items" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Result" canRead="True" canWrite="False" type="System.Int32" aliasName="Result" shouldSerialize="False" visibility="DefaultOn" source="" blockTypeName="" />
        </Items>
      </Content>
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="Add" />
            <MemberType Value="Method" />
            <TypeName Value="System.Int32" />
            <Content Name="Signature">
              <OpenSpan.Automation.MethodSignature>
                <ReturnType Value="System.Int32" />
                <Content Name="ParameterPrototype">
                  <Items>
                    <OpenSpan.Automation.ParameterPrototype>
                      <CanRead Value="False" />
                      <CanWrite Value="True" />
                      <DefaultSet Value="False" />
                      <ParamName Value="item" />
                      <Position Value="0" />
                      <TypeName Value="System.Object" />
                    </OpenSpan.Automation.ParameterPrototype>
                  </Items>
                </Content>
              </OpenSpan.Automation.MethodSignature>
            </Content>
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableMethod>
    <OpenSpan.Automation.ConnectableMethod Name="connectableMethod3" Id="ConnectableMethod-8D53AA3F4886930">
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
    <OpenSpan.Design.TypeProxy Name="_WorksheetProxy1" Id="TypeProxy-8D53AA40FCF2350">
      <AliasName Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel._Worksheet, Microsoft.Office.Interop.Excel" />
      <UseAlias Value="False" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Instance" canRead="True" canWrite="True" typeAssembly="Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" type="Microsoft.Office.Interop.Excel._Worksheet" aliasName="Instance" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Design.TypeProxy>
    <OpenSpan.Automation.ConnectableTypeProxy Name="connectableTypeProxy2" Id="ConnectableTypeProxy-8D53AA40FD3B730">
      <ComponentName Value="_WorksheetProxy1" />
      <DisplayName Value="Proxy" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel._Worksheet" />
      <InstanceUniqueId Value="Automator-8D53AA217ABD62C\TypeProxy-8D53AA40FCF2350" />
      <MemberDetails Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel._Worksheet" />
    </OpenSpan.Automation.ConnectableTypeProxy>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties3" Id="ConnectableProperties-8D53AA41B347740">
      <ComponentName Value="_WorksheetProxy1" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel._Worksheet" />
      <InstanceUniqueId Value="Automator-8D53AA217ABD62C\TypeProxy-8D53AA40FCF2350" />
      <MemberDetails Value=".Name Property" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="Name" />
            <MemberType Value="Property" />
            <TypeName Value="System.String" />
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableProperties>
  </Component>
</OpenSpanDesignDocument>