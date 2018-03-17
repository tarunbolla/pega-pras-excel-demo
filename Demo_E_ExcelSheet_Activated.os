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
    </AssemblyReferences>
    <DynamicAssemblyReferences />
    <FileReferences />
    <BuildReferences />
  </ComponentInfo>
  <Component Version="1.0">
    <OpenSpan.Automation.Automator Name="Demo_E_ExcelSheet_Activated" Id="Automator-8D53AA501660FD0">
      <AutomationDocument>
        <Name Value="" />
        <Size Value="5100, 5100" />
        <Objects>
          <ConnectionBlock>
            <DisplayName Value="MicrosoftExcel.SheetActivate" />
            <ConnectableUniqueId Value="Automator-8D53AA501660FD0\ConnectableEvent-8D53AA6755228B0" />
            <PartID Value="11" />
            <Left Value="40" />
            <Top Value="100" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="microsoftExcel1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock type="OpenSpan.Automation.Design.ConnectionBlocks.EntryPointExecuteBlock">
            <DisplayName Value="Execute" />
            <ConnectableUniqueId Value="Automator-8D53AA501660FD0\ConnectableMethod-8D53AA6834FB3B0" />
            <PartID Value="14" />
            <Left Value="440" />
            <Top Value="100" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="Demo_P_Cast_To_WorkSheet" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D53AA501660FD0\ConnectableProperties-8D53AA691E008D0" />
            <PartID Value="18" />
            <Left Value="880" />
            <Top Value="120" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="cmbTabs" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Proxy" />
            <ConnectableUniqueId Value="Automator-8D53AA501660FD0\ConnectableTypeProxy-8D53AA6A0276E10" />
            <PartID Value="22" />
            <Left Value="700" />
            <Top Value="280" />
            <Collapsed Value="False" />
            <WillExecute Value="False" />
            <InstanceName Value="_WorksheetProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D53AA501660FD0\ConnectableProperties-8D53AA6A3AE9900" />
            <PartID Value="24" />
            <Left Value="700" />
            <Top Value="200" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="_WorksheetProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D53AA501660FD0\ConnectableProperties-8D53AEAB0545BB0" />
            <PartID Value="26" />
            <Left Value="240" />
            <Top Value="100" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="microsoftExcel1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Proxy" />
            <ConnectableUniqueId Value="Automator-8D53AA501660FD0\ConnectableTypeProxy-8D53AEAB5690470" />
            <PartID Value="27" />
            <Left Value="220" />
            <Top Value="180" />
            <Collapsed Value="False" />
            <WillExecute Value="False" />
            <InstanceName Value="_WorkbookProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D53AA501660FD0\ConnectableProperties-8D53AEAD0EBB620" />
            <PartID Value="29" />
            <Left Value="220" />
            <Top Value="260" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="_WorkbookProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
        </Objects>
        <Links>
          <Link PartID="20" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="14" PortName="Complete" PortType="Event" ConnectableId="Automator-8D53AA501660FD0\ConnectableMethod-8D53AA6834FB3B0" MemberComponentId="Automator-8D53AA3CD4DFDD0\ExitPoint-8D53AA3DB396060" />
            <To PartID="18" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AA501660FD0\ConnectableProperties-8D53AA691E008D0" MemberComponentId="Automator-8D53AA501660FD0\ConnectableProperties-8D53AA691E008D0" />
            <LinkPoints>
              <Point value="677, 146" />
              <Point value="687, 146" />
              <Point value="687, 149" />
              <Point value="687, 149" />
              <Point value="875, 149" />
              <Point value="885, 149" />
            </LinkPoints>
          </Link>
          <Link PartID="23" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="14" PortName="Result" PortType="Property" ConnectableId="Automator-8D53AA501660FD0\ConnectableMethod-8D53AA6834FB3B0" MemberComponentId="Automator-8D53AA501660FD0\ConnectableMethod-8D53AA6834FB3B0" />
            <To PartID="22" PortName="Instance" PortType="Property" ConnectableId="Automator-8D53AA501660FD0\ConnectableTypeProxy-8D53AA6A0276E10" MemberComponentId="Automator-8D53AA501660FD0\TypeProxy-8D53AA6A0239D80" />
            <LinkPoints>
              <Point value="677, 180" />
              <Point value="687, 180" />
              <Point value="692, 180" />
              <Point value="692, 325" />
              <Point value="695, 325" />
              <Point value="705, 325" />
            </LinkPoints>
          </Link>
          <Link PartID="25" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="24" PortName="Name" PortType="Property" ConnectableId="Automator-8D53AA501660FD0\ConnectableProperties-8D53AA6A3AE9900" MemberComponentId="Automator-8D53AA501660FD0\TypeProxy-8D53AA6A0239D80" />
            <To PartID="18" PortName="Text" PortType="Property" ConnectableId="Automator-8D53AA501660FD0\ConnectableProperties-8D53AA691E008D0" MemberComponentId="DesignForm-8D53AA1BDDF1678\ComboBox-8D53AA1EEEFDA5E" />
            <LinkPoints>
              <Point value="865, 246" />
              <Point value="875, 246" />
              <Point value="876, 246" />
              <Point value="876, 166" />
              <Point value="875, 166" />
              <Point value="885, 166" />
            </LinkPoints>
          </Link>
          <Link PartID="28" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="26" PortName="ExcelWorkbook" PortType="Property" ConnectableId="Automator-8D53AA501660FD0\ConnectableProperties-8D53AEAB0545BB0" MemberComponentId="GlobalContainer-8D53AA1A011B17D\MicrosoftExcel-8D53AA1A49B9D1B" />
            <To PartID="27" PortName="Instance" PortType="Property" ConnectableId="Automator-8D53AA501660FD0\ConnectableTypeProxy-8D53AEAB5690470" MemberComponentId="Automator-8D53AA501660FD0\TypeProxy-8D53AEAB56533E0" />
            <LinkPoints>
              <Point value="386, 146" />
              <Point value="396, 146" />
              <Point value="396, 146" />
              <Point value="396, 164" />
              <Point value="212, 164" />
              <Point value="212, 225" />
              <Point value="215, 225" />
              <Point value="225, 225" />
            </LinkPoints>
          </Link>
          <Link PartID="30" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="29" PortName="ActiveSheet" PortType="Property" ConnectableId="Automator-8D53AA501660FD0\ConnectableProperties-8D53AEAD0EBB620" MemberComponentId="Automator-8D53AA501660FD0\TypeProxy-8D53AEAB56533E0" />
            <To PartID="14" PortName="param1" PortType="Property" ConnectableId="Automator-8D53AA501660FD0\ConnectableMethod-8D53AA6834FB3B0" MemberComponentId="Automator-8D53AA501660FD0\ConnectableMethod-8D53AA6834FB3B0" />
            <LinkPoints>
              <Point value="381, 306" />
              <Point value="391, 306" />
              <Point value="396, 306" />
              <Point value="396, 163" />
              <Point value="435, 163" />
              <Point value="445, 163" />
            </LinkPoints>
          </Link>
          <Link PartID="31" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="11" PortName="Raised" PortType="Event" ConnectableId="Automator-8D53AA501660FD0\ConnectableEvent-8D53AA6755228B0" MemberComponentId="Automator-8D53AA501660FD0\ConnectableEvent-8D53AA6755228B0" />
            <To PartID="26" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AA501660FD0\ConnectableProperties-8D53AEAB0545BB0" MemberComponentId="Automator-8D53AA501660FD0\ConnectableProperties-8D53AEAB0545BB0" />
            <LinkPoints>
              <Point value="186, 129" />
              <Point value="196, 129" />
              <Point value="196, 129" />
              <Point value="196, 129" />
              <Point value="235, 129" />
              <Point value="245, 129" />
            </LinkPoints>
          </Link>
          <Link PartID="34" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="26" PortName="Complete" PortType="Event" ConnectableId="Automator-8D53AA501660FD0\ConnectableProperties-8D53AEAB0545BB0" MemberComponentId="Automator-8D53AA501660FD0\ConnectableProperties-8D53AEAB0545BB0" />
            <To PartID="14" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AA501660FD0\ConnectableMethod-8D53AA6834FB3B0" MemberComponentId="Automator-8D53AA501660FD0\ConnectableMethod-8D53AA6834FB3B0" />
            <LinkPoints>
              <Point value="386, 129" />
              <Point value="396, 129" />
              <Point value="396, 129" />
              <Point value="396, 129" />
              <Point value="435, 129" />
              <Point value="445, 129" />
            </LinkPoints>
          </Link>
        </Links>
        <Comments />
        <SubGraphs />
      </AutomationDocument>
      <DocumentPosition Value="Binary">
        <Binary>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj00LjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5Qb2ludEYCAAAAAXgBeQAACwsCAAAAAAC2QwAAAAAL</Binary>
      </DocumentPosition>
    </OpenSpan.Automation.Automator>
    <OpenSpan.Automation.ConnectableEvent Name="connectableEvent2" Id="ConnectableEvent-8D53AA6755228B0">
      <ComponentName Value="microsoftExcel1" />
      <DisplayName Value="MicrosoftExcel.SheetActivate" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Office.MicrosoftExcel" />
      <InstanceUniqueId Value="GlobalContainer-8D53AA1A011B17D\MicrosoftExcel-8D53AA1A49B9D1B" />
      <MemberDetails Value=".SheetActivate Event" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="SheetActivate" />
            <MemberType Value="Event" />
            <TypeName Value="System.EventHandler" />
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableEvent>
    <OpenSpan.Automation.ConnectableMethod Name="connectableMethod1" Id="ConnectableMethod-8D53AA6834FB3B0">
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
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties1" Id="ConnectableProperties-8D53AA691E008D0">
      <ComponentName Value="cmbTabs" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="System.Windows.Forms.ComboBox" />
      <InstanceUniqueId Value="DesignForm-8D53AA1BDDF1678\ComboBox-8D53AA1EEEFDA5E" />
      <MemberDetails Value=".Text Property" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="Text" />
            <MemberType Value="Property" />
            <TypeName Value="System.String" />
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableProperties>
    <OpenSpan.Design.TypeProxy Name="_WorksheetProxy1" Id="TypeProxy-8D53AA6A0239D80">
      <AliasName Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel._Worksheet, Microsoft.Office.Interop.Excel" />
      <UseAlias Value="False" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Instance" canRead="True" canWrite="True" typeAssembly="Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" type="Microsoft.Office.Interop.Excel._Worksheet" aliasName="Instance" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Design.TypeProxy>
    <OpenSpan.Automation.ConnectableTypeProxy Name="connectableTypeProxy2" Id="ConnectableTypeProxy-8D53AA6A0276E10">
      <ComponentName Value="_WorksheetProxy1" />
      <DisplayName Value="Proxy" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel._Worksheet" />
      <InstanceUniqueId Value="Automator-8D53AA501660FD0\TypeProxy-8D53AA6A0239D80" />
      <MemberDetails Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel._Worksheet" />
    </OpenSpan.Automation.ConnectableTypeProxy>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties2" Id="ConnectableProperties-8D53AA6A3AE9900">
      <ComponentName Value="_WorksheetProxy1" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel._Worksheet" />
      <InstanceUniqueId Value="Automator-8D53AA501660FD0\TypeProxy-8D53AA6A0239D80" />
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
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties3" Id="ConnectableProperties-8D53AEAB0545BB0">
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
    <OpenSpan.Design.TypeProxy Name="_WorkbookProxy1" Id="TypeProxy-8D53AEAB56533E0">
      <AliasName Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel._Workbook, Microsoft.Office.Interop.Excel" />
      <UseAlias Value="False" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Instance" canRead="True" canWrite="True" typeAssembly="Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" type="Microsoft.Office.Interop.Excel._Workbook" aliasName="Instance" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Design.TypeProxy>
    <OpenSpan.Automation.ConnectableTypeProxy Name="connectableTypeProxy3" Id="ConnectableTypeProxy-8D53AEAB5690470">
      <ComponentName Value="_WorkbookProxy1" />
      <DisplayName Value="Proxy" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel._Workbook" />
      <InstanceUniqueId Value="Automator-8D53AA501660FD0\TypeProxy-8D53AEAB56533E0" />
      <MemberDetails Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel._Workbook" />
    </OpenSpan.Automation.ConnectableTypeProxy>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties4" Id="ConnectableProperties-8D53AEAD0EBB620">
      <ComponentName Value="_WorkbookProxy1" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel._Workbook" />
      <InstanceUniqueId Value="Automator-8D53AA501660FD0\TypeProxy-8D53AEAB56533E0" />
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
  </Component>
</OpenSpanDesignDocument>