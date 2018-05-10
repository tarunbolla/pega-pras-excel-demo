<OpenSpanDesignDocument Version="8.0.1026.0" Serializer="2.0" Culture="en-US">
  <ComponentInfo>
    <Type Value="OpenSpan.Automation.Automator" />
    <Assembly Value="OpenSpan.Automation" />
    <AssemblyReferences>
      <Assembly Value="mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" />
      <Assembly Value="System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" />
      <Assembly Value="System.Data, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" />
      <Assembly Value="System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" />
      <Assembly Value="Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" />
      <Assembly Value="OpenSpan, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" />
      <Assembly Value="OpenSpan.Automation, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" />
      <Assembly Value="OpenSpan.Controls, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" />
      <Assembly Value="OpenSpan.Office, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" />
      <Assembly Value="OpenSpan.Script, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" />
    </AssemblyReferences>
    <DynamicAssemblyReferences />
    <FileReferences />
    <BuildReferences />
  </ComponentInfo>
  <Component Version="1.0">
    <OpenSpan.Automation.Automator Name="Demo_P_GetUsedRangeForSheet" Id="Automator-8D540A8F1FCACEA">
      <AutomationDocument>
        <Name Value="" />
        <Size Value="5000, 5000" />
        <Objects>
          <ConnectionBlock>
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\EntryPoint-8D540AADD06B31A" />
            <Left Value="86" />
            <Top Value="104" />
            <PartID Value="1" />
          </ConnectionBlock>
          <ConnectionBlock type="OpenSpan.Automation.Design.ConnectionBlocks.EntryPointExecuteBlock">
            <DisplayName Value="Execute" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableMethod-8D540AAE4C51F1A" />
            <PartID Value="2" />
            <Left Value="260" />
            <Top Value="100" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="Demo_P_GetSheetByName" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Proxy" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableTypeProxy-8D540AAEF82251A" />
            <PartID Value="5" />
            <Left Value="320" />
            <Top Value="240" />
            <Collapsed Value="False" />
            <WillExecute Value="False" />
            <InstanceName Value="_WorksheetProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableProperties-8D540AB1682469A" />
            <PartID Value="8" />
            <Left Value="540" />
            <Top Value="120" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="_WorksheetProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Proxy" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableTypeProxy-8D540B3D107EF3A" />
            <PartID Value="29" />
            <Left Value="540" />
            <Top Value="240" />
            <Collapsed Value="False" />
            <WillExecute Value="False" />
            <InstanceName Value="rangeProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Select" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableMethod-8D540B3DC91BB1A" />
            <PartID Value="31" />
            <Left Value="820" />
            <Top Value="120" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="rangeProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65D0DADF834" />
            <PartID Value="33" />
            <Left Value="1000" />
            <Top Value="120" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="rangeProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Proxy" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableTypeProxy-8D5B65D143101A0" />
            <PartID Value="35" />
            <Left Value="1000" />
            <Top Value="200" />
            <Collapsed Value="False" />
            <WillExecute Value="False" />
            <InstanceName Value="RowsProxy" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65DBFCC3F97" />
            <PartID Value="37" />
            <Left Value="1780" />
            <Top Value="220" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="RowsProxy" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Concat" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E04D697B2" />
            <PartID Value="41" />
            <Left Value="1220" />
            <Top Value="120" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="stringUtils1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65E4893A6C5" />
            <PartID Value="43" />
            <Left Value="1560" />
            <Top Value="120" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="rangeProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Proxy" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableTypeProxy-8D5B65E574A81F1" />
            <PartID Value="45" />
            <Left Value="1560" />
            <Top Value="200" />
            <Collapsed Value="False" />
            <WillExecute Value="False" />
            <InstanceName Value="ColsProxy" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65E692790F0" />
            <PartID Value="47" />
            <Left Value="1560" />
            <Top Value="280" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="ColsProxy" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Concat" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E71A0F8F8" />
            <PartID Value="48" />
            <Left Value="2020" />
            <Top Value="120" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="stringUtils1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock type="OpenSpan.Automation.Design.ConnectionBlocks.EntryPointExecuteBlock">
            <DisplayName Value="Execute" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E8093CD3D" />
            <PartID Value="51" />
            <Left Value="2220" />
            <Top Value="120" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="Demo_P_ReturnDataTableForRange" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65E9E6FD898" />
            <PartID Value="53" />
            <Left Value="1400" />
            <Top Value="120" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="startCell" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65EC46BDF11" />
            <PartID Value="58" />
            <Left Value="2020" />
            <Top Value="240" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="startCell" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65EDCCB5B38" />
            <PartID Value="60" />
            <Left Value="2560" />
            <Top Value="120" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="dataSheet" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="NumberToCell" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B66116C76F8C" />
            <PartID Value="72" />
            <Left Value="1780" />
            <Top Value="120" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="HelperScripts" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Proxy" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableTypeProxy-8D5B6618A9FDFBD" />
            <PartID Value="79" />
            <Left Value="2220" />
            <Top Value="240" />
            <Collapsed Value="False" />
            <WillExecute Value="False" />
            <InstanceName Value="dataTableProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B661A969AC8E" />
            <PartID Value="81" />
            <Left Value="2720" />
            <Top Value="120" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="dataTableProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="StartLoop" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ListLoop-8D5B661B2C479AB" />
            <PartID Value="83" />
            <Left Value="2940" />
            <Top Value="120" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="listLoop1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock type="OpenSpan.Automation.Design.ConnectionBlocks.EntryPointExecuteBlock">
            <DisplayName Value="Execute" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B6625BF166D4" />
            <PartID Value="88" />
            <Left Value="3140" />
            <Top Value="120" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="Demo_P_Cast_To_DataRow" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Proxy" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableTypeProxy-8D5B6626732DAB6" />
            <PartID Value="91" />
            <Left Value="3140" />
            <Top Value="240" />
            <Collapsed Value="False" />
            <WillExecute Value="False" />
            <InstanceName Value="dataRowProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="SetItem" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B662825A1495" />
            <PartID Value="93" />
            <Left Value="3460" />
            <Top Value="140" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="dataRowProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B6629D9F3778" />
            <PartID Value="95" />
            <Left Value="3320" />
            <Top Value="240" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="startCell" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="ImportData" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B6630986ACFF" />
            <PartID Value="97" />
            <Left Value="3660" />
            <Top Value="340" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="microsoftExcel1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B66319EEECC3" />
            <PartID Value="99" />
            <Left Value="3460" />
            <Top Value="440" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="dataTableProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Close" />
            <ConnectableUniqueId Value="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B6638DBB9CD7" />
            <PartID Value="101" />
            <Left Value="3900" />
            <Top Value="340" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="microsoftExcel1" />
            <OverriddenIds />
          </ConnectionBlock>
        </Objects>
        <Links>
          <Link PartID="3" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="1" PortName="Complete" PortType="Event" ConnectableId="Automator-8D540A8F1FCACEA\EntryPoint-8D540AADD06B31A" MemberComponentId="Automator-8D540A8F1FCACEA\EntryPoint-8D540AADD06B31A" />
            <To PartID="2" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D540AAE4C51F1A" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D540AAE4C51F1A" />
            <LinkPoints>
              <Point value="197, 118" />
              <Point value="207, 118" />
              <Point value="207, 118" />
              <Point value="207, 129" />
              <Point value="255, 129" />
              <Point value="265, 129" />
            </LinkPoints>
          </Link>
          <Link PartID="4" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="1" PortName="param1" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\EntryPoint-8D540AADD06B31A" MemberComponentId="EMPTY" />
            <To PartID="2" PortName="_param1" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D540AAE4C51F1A" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D540AAE4C51F1A" />
            <LinkPoints>
              <Point value="197, 145" />
              <Point value="207, 145" />
              <Point value="207, 145" />
              <Point value="207, 163" />
              <Point value="255, 163" />
              <Point value="265, 163" />
            </LinkPoints>
          </Link>
          <Link PartID="6" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="2" PortName="Result" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D540AAE4C51F1A" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D540AAE4C51F1A" />
            <To PartID="5" PortName="Instance" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableTypeProxy-8D540AAEF82251A" MemberComponentId="Automator-8D540A8F1FCACEA\TypeProxy-8D540AAEF7C7FCA" />
            <LinkPoints>
              <Point value="480, 180" />
              <Point value="490, 180" />
              <Point value="492, 180" />
              <Point value="492, 196" />
              <Point value="316, 196" />
              <Point value="316, 285" />
              <Point value="315, 285" />
              <Point value="325, 285" />
            </LinkPoints>
          </Link>
          <Link PartID="24" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="2" PortName="Complete" PortType="Event" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D540AAE4C51F1A" MemberComponentId="Automator-8D53AEC9D8E3A30\ExitPoint-8D53AEE897AB120" />
            <To PartID="8" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D540AB1682469A" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D540AB1682469A" />
            <LinkPoints>
              <Point value="480, 146" />
              <Point value="490, 146" />
              <Point value="513, 146" />
              <Point value="513, 149" />
              <Point value="535, 149" />
              <Point value="545, 149" />
            </LinkPoints>
          </Link>
          <Link PartID="30" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="8" PortName="UsedRange" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D540AB1682469A" MemberComponentId="Automator-8D540A8F1FCACEA\TypeProxy-8D540AAEF7C7FCA" />
            <To PartID="29" PortName="Instance" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableTypeProxy-8D540B3D107EF3A" MemberComponentId="Automator-8D540A8F1FCACEA\TypeProxy-8D540B3D1041EAA" />
            <LinkPoints>
              <Point value="705, 166" />
              <Point value="715, 166" />
              <Point value="716, 166" />
              <Point value="716, 180" />
              <Point value="532, 180" />
              <Point value="532, 285" />
              <Point value="535, 285" />
              <Point value="545, 285" />
            </LinkPoints>
          </Link>
          <Link PartID="32" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="8" PortName="Complete" PortType="Event" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D540AB1682469A" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D540AB1682469A" />
            <To PartID="31" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D540B3DC91BB1A" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D540B3DC91BB1A" />
            <LinkPoints>
              <Point value="705, 149" />
              <Point value="715, 149" />
              <Point value="765, 149" />
              <Point value="765, 149" />
              <Point value="815, 149" />
              <Point value="825, 149" />
            </LinkPoints>
          </Link>
          <Link PartID="34" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="31" PortName="Complete" PortType="Event" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D540B3DC91BB1A" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D540B3DC91BB1A" />
            <To PartID="33" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65D0DADF834" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65D0DADF834" />
            <LinkPoints>
              <Point value="946, 149" />
              <Point value="956, 149" />
              <Point value="976, 149" />
              <Point value="976, 149" />
              <Point value="995, 149" />
              <Point value="1005, 149" />
            </LinkPoints>
          </Link>
          <Link PartID="36" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="33" PortName="Rows" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65D0DADF834" MemberComponentId="Automator-8D540A8F1FCACEA\TypeProxy-8D540B3D1041EAA" />
            <To PartID="35" PortName="Instance" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableTypeProxy-8D5B65D143101A0" MemberComponentId="Automator-8D540A8F1FCACEA\TypeProxy-8D5B65D141C4098" />
            <LinkPoints>
              <Point value="1126, 166" />
              <Point value="1136, 166" />
              <Point value="1140, 166" />
              <Point value="1140, 180" />
              <Point value="996, 180" />
              <Point value="996, 245" />
              <Point value="995, 245" />
              <Point value="1005, 245" />
            </LinkPoints>
          </Link>
          <Link PartID="42" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="33" PortName="Complete" PortType="Event" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65D0DADF834" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65D0DADF834" />
            <To PartID="41" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E04D697B2" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E04D697B2" />
            <LinkPoints>
              <Point value="1126, 149" />
              <Point value="1136, 149" />
              <Point value="1136, 149" />
              <Point value="1136, 149" />
              <Point value="1215, 149" />
              <Point value="1225, 149" />
            </LinkPoints>
          </Link>
          <Link PartID="46" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="43" PortName="Columns" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65E4893A6C5" MemberComponentId="Automator-8D540A8F1FCACEA\TypeProxy-8D540B3D1041EAA" />
            <To PartID="45" PortName="Instance" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableTypeProxy-8D5B65E574A81F1" MemberComponentId="Automator-8D540A8F1FCACEA\TypeProxy-8D5B65E57474D8C" />
            <LinkPoints>
              <Point value="1686, 166" />
              <Point value="1696, 166" />
              <Point value="1700, 166" />
              <Point value="1700, 188" />
              <Point value="1556, 188" />
              <Point value="1556, 245" />
              <Point value="1555, 245" />
              <Point value="1565, 245" />
            </LinkPoints>
          </Link>
          <Link PartID="52" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="48" PortName="Result" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E71A0F8F8" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E71A0F8F8" />
            <To PartID="51" PortName="param2" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E8093CD3D" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E8093CD3D" />
            <LinkPoints>
              <Point value="2137, 217" />
              <Point value="2147, 217" />
              <Point value="2148, 217" />
              <Point value="2148, 200" />
              <Point value="2215, 200" />
              <Point value="2225, 200" />
            </LinkPoints>
          </Link>
          <Link PartID="54" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="41" PortName="Result" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E04D697B2" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E04D697B2" />
            <To PartID="53" PortName="Value" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65E9E6FD898" MemberComponentId="Automator-8D540A8F1FCACEA\StringVariable-8D5B65E8FC2E8BD" />
            <LinkPoints>
              <Point value="1337, 200" />
              <Point value="1347, 200" />
              <Point value="1348, 200" />
              <Point value="1348, 166" />
              <Point value="1395, 166" />
              <Point value="1405, 166" />
            </LinkPoints>
          </Link>
          <Link PartID="55" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="41" PortName="Complete" PortType="Event" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E04D697B2" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E04D697B2" />
            <To PartID="53" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65E9E6FD898" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65E9E6FD898" />
            <LinkPoints>
              <Point value="1337, 149" />
              <Point value="1347, 149" />
              <Point value="1371, 149" />
              <Point value="1371, 149" />
              <Point value="1395, 149" />
              <Point value="1405, 149" />
            </LinkPoints>
          </Link>
          <Link PartID="56" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="53" PortName="Complete" PortType="Event" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65E9E6FD898" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65E9E6FD898" />
            <To PartID="43" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65E4893A6C5" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65E4893A6C5" />
            <LinkPoints>
              <Point value="1509, 149" />
              <Point value="1519, 149" />
              <Point value="1537, 149" />
              <Point value="1537, 149" />
              <Point value="1555, 149" />
              <Point value="1565, 149" />
            </LinkPoints>
          </Link>
          <Link PartID="57" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="48" PortName="Complete" PortType="Event" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E71A0F8F8" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E71A0F8F8" />
            <To PartID="51" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E8093CD3D" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E8093CD3D" />
            <LinkPoints>
              <Point value="2137, 149" />
              <Point value="2147, 149" />
              <Point value="2181, 149" />
              <Point value="2181, 149" />
              <Point value="2215, 149" />
              <Point value="2225, 149" />
            </LinkPoints>
          </Link>
          <Link PartID="59" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="58" PortName="Value" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65EC46BDF11" MemberComponentId="Automator-8D540A8F1FCACEA\StringVariable-8D5B65E8FC2E8BD" />
            <To PartID="51" PortName="param1" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E8093CD3D" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E8093CD3D" />
            <LinkPoints>
              <Point value="2129, 286" />
              <Point value="2139, 286" />
              <Point value="2140, 286" />
              <Point value="2140, 286" />
              <Point value="2148, 286" />
              <Point value="2148, 183" />
              <Point value="2215, 183" />
              <Point value="2225, 183" />
            </LinkPoints>
          </Link>
          <Link PartID="61" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="51" PortName="Complete" PortType="Event" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E8093CD3D" MemberComponentId="Automator-8D53B34D7853250\ExitPoint-8D53B34EDF49AD0" />
            <To PartID="60" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65EDCCB5B38" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65EDCCB5B38" />
            <LinkPoints>
              <Point value="2496, 166" />
              <Point value="2506, 166" />
              <Point value="2530, 166" />
              <Point value="2530, 149" />
              <Point value="2555, 149" />
              <Point value="2565, 149" />
            </LinkPoints>
          </Link>
          <Link PartID="62" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="51" PortName="Result" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E8093CD3D" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E8093CD3D" />
            <To PartID="60" PortName="DataSource" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65EDCCB5B38" MemberComponentId="DesignForm-8D53AA1BDDF1678\DataGridView-8D53AA1F6ACCB63" />
            <LinkPoints>
              <Point value="2496, 217" />
              <Point value="2506, 217" />
              <Point value="2508, 217" />
              <Point value="2508, 166" />
              <Point value="2555, 166" />
              <Point value="2565, 166" />
            </LinkPoints>
          </Link>
          <Link PartID="73" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="43" PortName="Complete" PortType="Event" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65E4893A6C5" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65E4893A6C5" />
            <To PartID="72" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B66116C76F8C" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B66116C76F8C" />
            <LinkPoints>
              <Point value="1686, 149" />
              <Point value="1696, 149" />
              <Point value="1696, 149" />
              <Point value="1696, 149" />
              <Point value="1775, 149" />
              <Point value="1785, 149" />
            </LinkPoints>
          </Link>
          <Link PartID="74" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="72" PortName="Complete" PortType="Event" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B66116C76F8C" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B66116C76F8C" />
            <To PartID="48" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E71A0F8F8" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E71A0F8F8" />
            <LinkPoints>
              <Point value="1914, 149" />
              <Point value="1924, 149" />
              <Point value="1924, 149" />
              <Point value="1924, 149" />
              <Point value="2015, 149" />
              <Point value="2025, 149" />
            </LinkPoints>
          </Link>
          <Link PartID="75" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="47" PortName="Count" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65E692790F0" MemberComponentId="Automator-8D540A8F1FCACEA\TypeProxy-8D5B65E57474D8C" />
            <To PartID="72" PortName="number" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B66116C76F8C" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B66116C76F8C" />
            <LinkPoints>
              <Point value="1671, 326" />
              <Point value="1681, 326" />
              <Point value="1684, 326" />
              <Point value="1684, 326" />
              <Point value="1700, 326" />
              <Point value="1700, 166" />
              <Point value="1775, 166" />
              <Point value="1785, 166" />
            </LinkPoints>
          </Link>
          <Link PartID="77" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="72" PortName="Result" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B66116C76F8C" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B66116C76F8C" />
            <To PartID="48" PortName="list0" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E71A0F8F8" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E71A0F8F8" />
            <LinkPoints>
              <Point value="1914, 183" />
              <Point value="1924, 183" />
              <Point value="1924, 183" />
              <Point value="1924, 166" />
              <Point value="2015, 166" />
              <Point value="2025, 166" />
            </LinkPoints>
          </Link>
          <Link PartID="78" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="37" PortName="Count" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65DBFCC3F97" MemberComponentId="Automator-8D540A8F1FCACEA\TypeProxy-8D5B65D141C4098" />
            <To PartID="48" PortName="list1" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E71A0F8F8" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E71A0F8F8" />
            <LinkPoints>
              <Point value="1898, 266" />
              <Point value="1908, 266" />
              <Point value="1908, 266" />
              <Point value="1908, 266" />
              <Point value="1924, 266" />
              <Point value="1924, 183" />
              <Point value="2015, 183" />
              <Point value="2025, 183" />
            </LinkPoints>
          </Link>
          <Link PartID="80" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="51" PortName="Result" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E8093CD3D" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B65E8093CD3D" />
            <To PartID="79" PortName="Instance" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableTypeProxy-8D5B6618A9FDFBD" MemberComponentId="Automator-8D540A8F1FCACEA\TypeProxy-8D5B6618A9CAB58" />
            <LinkPoints>
              <Point value="2496, 217" />
              <Point value="2506, 217" />
              <Point value="2508, 217" />
              <Point value="2508, 236" />
              <Point value="2212, 236" />
              <Point value="2212, 285" />
              <Point value="2215, 285" />
              <Point value="2225, 285" />
            </LinkPoints>
          </Link>
          <Link PartID="82" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="60" PortName="Complete" PortType="Event" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65EDCCB5B38" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B65EDCCB5B38" />
            <To PartID="81" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B661A969AC8E" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B661A969AC8E" />
            <LinkPoints>
              <Point value="2671, 149" />
              <Point value="2681, 149" />
              <Point value="2698, 149" />
              <Point value="2698, 149" />
              <Point value="2715, 149" />
              <Point value="2725, 149" />
            </LinkPoints>
          </Link>
          <Link PartID="84" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="81" PortName="Complete" PortType="Event" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B661A969AC8E" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B661A969AC8E" />
            <To PartID="83" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D540A8F1FCACEA\ListLoop-8D5B661B2C479AB" MemberComponentId="Automator-8D540A8F1FCACEA\ListLoop-8D5B661B2C479AB" />
            <LinkPoints>
              <Point value="2873, 149" />
              <Point value="2883, 149" />
              <Point value="2909, 149" />
              <Point value="2909, 149" />
              <Point value="2935, 149" />
              <Point value="2945, 149" />
            </LinkPoints>
          </Link>
          <Link PartID="85" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="81" PortName="Rows" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B661A969AC8E" MemberComponentId="Automator-8D540A8F1FCACEA\TypeProxy-8D5B6618A9CAB58" />
            <To PartID="83" PortName="List" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ListLoop-8D5B661B2C479AB" MemberComponentId="Automator-8D540A8F1FCACEA\ListLoop-8D5B661B2C479AB" />
            <LinkPoints>
              <Point value="2873, 166" />
              <Point value="2883, 166" />
              <Point value="2909, 166" />
              <Point value="2909, 166" />
              <Point value="2935, 166" />
              <Point value="2945, 166" />
            </LinkPoints>
          </Link>
          <Link PartID="89" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="83" PortName="TypedItem" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ListLoop-8D5B661B2C479AB" MemberComponentId="Automator-8D540A8F1FCACEA\ListLoop-8D5B661B2C479AB" />
            <To PartID="88" PortName="param1" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B6625BF166D4" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B6625BF166D4" />
            <LinkPoints>
              <Point value="3059, 200" />
              <Point value="3069, 200" />
              <Point value="3069, 200" />
              <Point value="3069, 183" />
              <Point value="3135, 183" />
              <Point value="3145, 183" />
            </LinkPoints>
          </Link>
          <Link PartID="90" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="83" PortName="ItemYielded" PortType="Event" ConnectableId="Automator-8D540A8F1FCACEA\ListLoop-8D5B661B2C479AB" MemberComponentId="Automator-8D540A8F1FCACEA\ListLoop-8D5B661B2C479AB" />
            <To PartID="88" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B6625BF166D4" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B6625BF166D4" />
            <LinkPoints>
              <Point value="3059, 183" />
              <Point value="3069, 183" />
              <Point value="3102, 183" />
              <Point value="3102, 149" />
              <Point value="3135, 149" />
              <Point value="3145, 149" />
            </LinkPoints>
          </Link>
          <Link PartID="92" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="88" PortName="Result" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B6625BF166D4" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B6625BF166D4" />
            <To PartID="91" PortName="Instance" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableTypeProxy-8D5B6626732DAB6" MemberComponentId="Automator-8D540A8F1FCACEA\TypeProxy-8D5B662672FCD62" />
            <LinkPoints>
              <Point value="3365, 200" />
              <Point value="3375, 200" />
              <Point value="3380, 200" />
              <Point value="3380, 220" />
              <Point value="3132, 220" />
              <Point value="3132, 285" />
              <Point value="3135, 285" />
              <Point value="3145, 285" />
            </LinkPoints>
          </Link>
          <Link PartID="94" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="88" PortName="Complete" PortType="Event" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B6625BF166D4" MemberComponentId="Automator-8D5B661C6B58228\ExitPoint-8D5B66247290BA4" />
            <To PartID="93" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B662825A1495" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B662825A1495" />
            <LinkPoints>
              <Point value="3365, 166" />
              <Point value="3375, 166" />
              <Point value="3380, 166" />
              <Point value="3380, 169" />
              <Point value="3455, 169" />
              <Point value="3465, 169" />
            </LinkPoints>
          </Link>
          <Link PartID="96" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="95" PortName="Value" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B6629D9F3778" MemberComponentId="Automator-8D540A8F1FCACEA\StringVariable-8D5B65E8FC2E8BD" />
            <To PartID="93" PortName="value" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B662825A1495" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B662825A1495" />
            <LinkPoints>
              <Point value="3429, 286" />
              <Point value="3439, 286" />
              <Point value="3444, 286" />
              <Point value="3444, 203" />
              <Point value="3455, 203" />
              <Point value="3465, 203" />
            </LinkPoints>
          </Link>
          <Link PartID="98" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="83" PortName="Completed" PortType="Event" ConnectableId="Automator-8D540A8F1FCACEA\ListLoop-8D5B661B2C479AB" MemberComponentId="Automator-8D540A8F1FCACEA\ListLoop-8D5B661B2C479AB" />
            <To PartID="97" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B6630986ACFF" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B6630986ACFF" />
            <LinkPoints>
              <Point value="3059, 234" />
              <Point value="3069, 234" />
              <Point value="3069, 234" />
              <Point value="3069, 369" />
              <Point value="3655, 369" />
              <Point value="3665, 369" />
            </LinkPoints>
          </Link>
          <Link PartID="100" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="99" PortName="This" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableProperties-8D5B66319EEECC3" MemberComponentId="Automator-8D540A8F1FCACEA\TypeProxy-8D5B6618A9CAB58" />
            <To PartID="97" PortName="dataTable" PortType="Property" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B6630986ACFF" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B6630986ACFF" />
            <LinkPoints>
              <Point value="3613, 486" />
              <Point value="3623, 486" />
              <Point value="3628, 486" />
              <Point value="3628, 386" />
              <Point value="3655, 386" />
              <Point value="3665, 386" />
            </LinkPoints>
          </Link>
          <Link PartID="102" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="97" PortName="Complete" PortType="Event" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B6630986ACFF" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B6630986ACFF" />
            <To PartID="101" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B6638DBB9CD7" MemberComponentId="Automator-8D540A8F1FCACEA\ConnectableMethod-8D5B6638DBB9CD7" />
            <LinkPoints>
              <Point value="3806, 369" />
              <Point value="3816, 369" />
              <Point value="3856, 369" />
              <Point value="3856, 369" />
              <Point value="3895, 369" />
              <Point value="3905, 369" />
            </LinkPoints>
          </Link>
        </Links>
        <Comments />
        <SubGraphs />
      </AutomationDocument>
      <DocumentPosition Value="Binary">
        <Binary>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj00LjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5Qb2ludEYCAAAAAXgBeQAACwsCAAAAAABIRal08EIL</Binary>
      </DocumentPosition>
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicMethodInfo dynamicType="Method" name="_EntryPointExecute" aliasName="Execute" visibility="DefaultOn" source="" blockTypeName="OpenSpan.Automation.Design.ConnectionBlocks.EntryPointExecuteBlock" returnType="System.Void">
            <param name="param1" aliasName="sheetName" paramType="System.String" isIn="True" isOut="False" position="0" />
          </OpenSpan.DynamicMembers.DynamicMethodInfo>
        </Items>
      </Content>
    </OpenSpan.Automation.Automator>
    <OpenSpan.Automation.EntryPoint Name="entryPoint1" Id="EntryPoint-8D540AADD06B31A">
      <AliasName Value="Execute" />
      <ComponentName Value="Execute" />
      <DisplayName Value="Execute" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Automation.EntryPoint" />
      <InstanceUniqueId Value="Automator-8D540A8F1FCACEA\EntryPoint-8D540AADD06B31A" />
      <MemberDetails Value="" />
      <MethodName Value="_EntryPointExecute" />
      <Removing Value="False" />
      <UseAlias Value="True" />
      <Content Name="Controls">
        <Capacity Value="4" />
        <Items>
          <OpenSpan.Automation.HiddenTypeProxy Name="hiddenTypeProxy1" Id="HiddenTypeProxy-8D540AADE1F27FA">
            <AliasName Value="sheetName" />
            <Parent Value="ComponentReference" Name="entryPoint1" />
            <ProxiedTypeName Value="System.String, mscorlib" />
            <Scope Value="Local" Extended="True" />
            <UseAlias Value="True" />
            <Content Name="DynamicMembers">
              <Items>
                <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Instance" canRead="True" canWrite="True" type="System.String" aliasName="Instance" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
              </Items>
            </Content>
          </OpenSpan.Automation.HiddenTypeProxy>
        </Items>
      </Content>
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Result" canRead="True" canWrite="True" type="System.Void" aliasName="Result" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="param1" canRead="False" canWrite="True" type="System.String" aliasName="sheetName" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Automation.EntryPoint>
    <OpenSpan.Automation.ConnectableMethod Name="connectableMethod1" Id="ConnectableMethod-8D540AAE4C51F1A">
      <ComponentName Value="Demo_P_GetSheetByName" />
      <DisplayName Value="Execute" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Automation.Automator" />
      <InstanceUniqueId Value="Automator-8D53AEC9D8E3A30" />
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
                      <ParamName Value="_param1" />
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
    <OpenSpan.Design.TypeProxy Name="_WorksheetProxy1" Id="TypeProxy-8D540AAEF7C7FCA">
      <AliasName Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel._Worksheet, Microsoft.Office.Interop.Excel" />
      <UseAlias Value="False" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Instance" canRead="True" canWrite="True" typeAssembly="Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" type="Microsoft.Office.Interop.Excel._Worksheet" aliasName="Instance" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Design.TypeProxy>
    <OpenSpan.Automation.ConnectableTypeProxy Name="connectableTypeProxy1" Id="ConnectableTypeProxy-8D540AAEF82251A">
      <ComponentName Value="_WorksheetProxy1" />
      <DisplayName Value="Proxy" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel._Worksheet" />
      <InstanceUniqueId Value="Automator-8D540A8F1FCACEA\TypeProxy-8D540AAEF7C7FCA" />
      <MemberDetails Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel._Worksheet" />
    </OpenSpan.Automation.ConnectableTypeProxy>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties1" Id="ConnectableProperties-8D540AB1682469A">
      <ComponentName Value="_WorksheetProxy1" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel._Worksheet" />
      <InstanceUniqueId Value="Automator-8D540A8F1FCACEA\TypeProxy-8D540AAEF7C7FCA" />
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
    <OpenSpan.Design.TypeProxy Name="rangeProxy1" Id="TypeProxy-8D540B3D1041EAA">
      <AliasName Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel.Range, Microsoft.Office.Interop.Excel" />
      <UseAlias Value="False" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Instance" canRead="True" canWrite="True" typeAssembly="Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" type="Microsoft.Office.Interop.Excel.Range" aliasName="Instance" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Design.TypeProxy>
    <OpenSpan.Automation.ConnectableTypeProxy Name="connectableTypeProxy2" Id="ConnectableTypeProxy-8D540B3D107EF3A">
      <ComponentName Value="rangeProxy1" />
      <DisplayName Value="Proxy" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel.Range" />
      <InstanceUniqueId Value="Automator-8D540A8F1FCACEA\TypeProxy-8D540B3D1041EAA" />
      <MemberDetails Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel.Range" />
    </OpenSpan.Automation.ConnectableTypeProxy>
    <OpenSpan.Automation.ConnectableMethod Name="connectableMethod2" Id="ConnectableMethod-8D540B3DC91BB1A">
      <ComponentName Value="rangeProxy1" />
      <DisplayName Value="Select" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel.Range" />
      <InstanceUniqueId Value="Automator-8D540A8F1FCACEA\TypeProxy-8D540B3D1041EAA" />
      <MemberDetails Value=".Select() Method" />
      <ParamsLength Value="0" />
      <SerializedParamsDefaultValues Value="" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Result" canRead="True" canWrite="False" type="System.Object" aliasName="Result" shouldSerialize="False" visibility="DefaultOn" source="" blockTypeName="" />
        </Items>
      </Content>
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="Select" />
            <MemberType Value="Method" />
            <TypeName Value="System.Object" />
            <Content Name="Signature">
              <OpenSpan.Automation.MethodSignature>
                <ReturnType Value="System.Object" />
              </OpenSpan.Automation.MethodSignature>
            </Content>
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableMethod>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties2" Id="ConnectableProperties-8D5B65D0DADF834">
      <ComponentName Value="rangeProxy1" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel.Range" />
      <InstanceUniqueId Value="Automator-8D540A8F1FCACEA\TypeProxy-8D540B3D1041EAA" />
      <MemberDetails Value=".Rows Property" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="Rows" />
            <MemberType Value="Property" />
            <TypeAssemblyName Value="Microsoft.Office.Interop.Excel" />
            <TypeName Value="Microsoft.Office.Interop.Excel.Range" />
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableProperties>
    <OpenSpan.Design.TypeProxy Name="RowsProxy" Id="TypeProxy-8D5B65D141C4098">
      <AliasName Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel.Range, Microsoft.Office.Interop.Excel" />
      <UseAlias Value="False" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Instance" canRead="True" canWrite="True" typeAssembly="Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" type="Microsoft.Office.Interop.Excel.Range" aliasName="Instance" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Design.TypeProxy>
    <OpenSpan.Automation.ConnectableTypeProxy Name="connectableTypeProxy3" Id="ConnectableTypeProxy-8D5B65D143101A0">
      <ComponentName Value="RowsProxy" />
      <DisplayName Value="Proxy" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel.Range" />
      <InstanceUniqueId Value="Automator-8D540A8F1FCACEA\TypeProxy-8D5B65D141C4098" />
      <MemberDetails Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel.Range" />
    </OpenSpan.Automation.ConnectableTypeProxy>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties3" Id="ConnectableProperties-8D5B65DBFCC3F97">
      <ComponentName Value="RowsProxy" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel.Range" />
      <InstanceUniqueId Value="Automator-8D540A8F1FCACEA\TypeProxy-8D5B65D141C4098" />
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
    <OpenSpan.Controls.StringUtils Name="stringUtils1" Id="StringUtils-8D5B65DFD138C05" />
    <OpenSpan.Automation.ConnectableMethod Name="connectableMethod3" Id="ConnectableMethod-8D5B65E04D697B2">
      <ComponentName Value="stringUtils1" />
      <DisplayName Value="Concat" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Controls.StringUtils" />
      <InstanceUniqueId Value="Automator-8D540A8F1FCACEA\StringUtils-8D5B65DFD138C05" />
      <MemberDetails Value=".Concat() Method" />
      <ParamsDefaultValues>
        <list0 defaultValue="A1" />
      </ParamsDefaultValues>
      <ParamsLength Value="2" />
      <SerializedParamsDefaultValues Value="" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Result" canRead="True" canWrite="False" type="System.String" aliasName="Result" shouldSerialize="False" visibility="DefaultOn" source="" blockTypeName="" />
        </Items>
      </Content>
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="Concat" />
            <MemberType Value="Method" />
            <TypeName Value="System.String" />
            <Content Name="Signature">
              <OpenSpan.Automation.MethodSignature>
                <ReturnType Value="System.String" />
                <Content Name="ParameterPrototype">
                  <Items>
                    <OpenSpan.Automation.ParameterPrototype>
                      <CanRead Value="False" />
                      <CanWrite Value="True" />
                      <DefaultSet Value="False" />
                      <ParamName Value="list" />
                      <Position Value="0" />
                      <TypeName Value="System.String[]" />
                    </OpenSpan.Automation.ParameterPrototype>
                  </Items>
                </Content>
              </OpenSpan.Automation.MethodSignature>
            </Content>
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableMethod>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties4" Id="ConnectableProperties-8D5B65E4893A6C5">
      <ComponentName Value="rangeProxy1" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel.Range" />
      <InstanceUniqueId Value="Automator-8D540A8F1FCACEA\TypeProxy-8D540B3D1041EAA" />
      <MemberDetails Value=".Columns Property" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="Columns" />
            <MemberType Value="Property" />
            <TypeAssemblyName Value="Microsoft.Office.Interop.Excel" />
            <TypeName Value="Microsoft.Office.Interop.Excel.Range" />
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableProperties>
    <OpenSpan.Design.TypeProxy Name="ColsProxy" Id="TypeProxy-8D5B65E57474D8C">
      <AliasName Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel.Range, Microsoft.Office.Interop.Excel" />
      <UseAlias Value="False" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Instance" canRead="True" canWrite="True" typeAssembly="Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" type="Microsoft.Office.Interop.Excel.Range" aliasName="Instance" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Design.TypeProxy>
    <OpenSpan.Automation.ConnectableTypeProxy Name="connectableTypeProxy4" Id="ConnectableTypeProxy-8D5B65E574A81F1">
      <ComponentName Value="ColsProxy" />
      <DisplayName Value="Proxy" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel.Range" />
      <InstanceUniqueId Value="Automator-8D540A8F1FCACEA\TypeProxy-8D5B65E57474D8C" />
      <MemberDetails Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel.Range" />
    </OpenSpan.Automation.ConnectableTypeProxy>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties5" Id="ConnectableProperties-8D5B65E692790F0">
      <ComponentName Value="ColsProxy" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel.Range" />
      <InstanceUniqueId Value="Automator-8D540A8F1FCACEA\TypeProxy-8D5B65E57474D8C" />
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
    <OpenSpan.Automation.ConnectableMethod Name="connectableMethod4" Id="ConnectableMethod-8D5B65E71A0F8F8">
      <ComponentName Value="stringUtils1" />
      <DisplayName Value="Concat" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Controls.StringUtils" />
      <InstanceUniqueId Value="Automator-8D540A8F1FCACEA\StringUtils-8D5B65DFD138C05" />
      <MemberDetails Value=".Concat() Method" />
      <ParamsLength Value="3" />
      <SerializedParamsDefaultValues Value="" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Result" canRead="True" canWrite="False" type="System.String" aliasName="Result" shouldSerialize="False" visibility="DefaultOn" source="" blockTypeName="" />
        </Items>
      </Content>
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="Concat" />
            <MemberType Value="Method" />
            <TypeName Value="System.String" />
            <Content Name="Signature">
              <OpenSpan.Automation.MethodSignature>
                <ReturnType Value="System.String" />
                <Content Name="ParameterPrototype">
                  <Items>
                    <OpenSpan.Automation.ParameterPrototype>
                      <CanRead Value="False" />
                      <CanWrite Value="True" />
                      <DefaultSet Value="False" />
                      <ParamName Value="list" />
                      <Position Value="0" />
                      <TypeName Value="System.String[]" />
                    </OpenSpan.Automation.ParameterPrototype>
                  </Items>
                </Content>
              </OpenSpan.Automation.MethodSignature>
            </Content>
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableMethod>
    <OpenSpan.Automation.ConnectableMethod Name="connectableMethod5" Id="ConnectableMethod-8D5B65E8093CD3D">
      <ComponentName Value="Demo_P_ReturnDataTableForRange" />
      <DisplayName Value="Execute" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Automation.Automator" />
      <InstanceUniqueId Value="Automator-8D53B34D7853250" />
      <MemberDetails Value=".Execute() Method" />
      <ParamsLength Value="0" />
      <SerializedParamsDefaultValues Value="" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Result" canRead="True" canWrite="False" type="System.Data.DataTable" aliasName="Result" shouldSerialize="False" visibility="DefaultOn" source="" blockTypeName="" />
        </Items>
      </Content>
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="_EntryPointExecute" />
            <MemberType Value="Method" />
            <TypeAssemblyName Value="System.Data" />
            <TypeName Value="System.Data.DataTable" />
            <Content Name="Signature">
              <OpenSpan.Automation.MethodSignature>
                <ReturnType Value="System.Data.DataTable" />
                <Content Name="ParameterPrototype">
                  <Items>
                    <OpenSpan.Automation.ParameterPrototype>
                      <CanRead Value="False" />
                      <CanWrite Value="True" />
                      <DefaultSet Value="False" />
                      <ParamName Value="param1" />
                      <Position Value="0" />
                      <TypeName Value="System.String" />
                    </OpenSpan.Automation.ParameterPrototype>
                    <OpenSpan.Automation.ParameterPrototype>
                      <CanRead Value="False" />
                      <CanWrite Value="True" />
                      <DefaultSet Value="False" />
                      <ParamName Value="param2" />
                      <Position Value="1" />
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
    <OpenSpan.Controls.StringVariable Name="startCell" Id="StringVariable-8D5B65E8FC2E8BD">
      <Value Value="" />
    </OpenSpan.Controls.StringVariable>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties6" Id="ConnectableProperties-8D5B65E9E6FD898">
      <ComponentName Value="startCell" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Controls.StringVariable" />
      <InstanceUniqueId Value="Automator-8D540A8F1FCACEA\StringVariable-8D5B65E8FC2E8BD" />
      <MemberDetails Value=".Value Property" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="Value" />
            <MemberType Value="Property" />
            <TypeName Value="System.String" />
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableProperties>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties7" Id="ConnectableProperties-8D5B65EC46BDF11">
      <ComponentName Value="startCell" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Controls.StringVariable" />
      <InstanceUniqueId Value="Automator-8D540A8F1FCACEA\StringVariable-8D5B65E8FC2E8BD" />
      <MemberDetails Value=".Value Property" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="Value" />
            <MemberType Value="Property" />
            <TypeName Value="System.String" />
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableProperties>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties8" Id="ConnectableProperties-8D5B65EDCCB5B38">
      <ComponentName Value="dataSheet" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="System.Windows.Forms.DataGridView" />
      <InstanceUniqueId Value="DesignForm-8D53AA1BDDF1678\DataGridView-8D53AA1F6ACCB63" />
      <MemberDetails Value=".DataSource Property" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="DataSource" />
            <MemberType Value="Property" />
            <TypeName Value="System.Object" />
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableProperties>
    <OpenSpan.Automation.ConnectableMethod Name="connectableMethod6" Id="ConnectableMethod-8D5B66116C76F8C">
      <ComponentName Value="HelperScripts" />
      <DisplayName Value="NumberToCell" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Script.Custom.Script" />
      <InstanceUniqueId Value="GlobalContainer-8D53AA1A011B17D\Script-8D5B65FEC42AD8F" />
      <MemberDetails Value=".NumberToCell() Method" />
      <ParamsLength Value="0" />
      <SerializedParamsDefaultValues Value="" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Result" canRead="True" canWrite="False" type="System.String" aliasName="Result" shouldSerialize="False" visibility="DefaultOn" source="" blockTypeName="" />
        </Items>
      </Content>
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="NumberToCell" />
            <MemberType Value="Method" />
            <TypeName Value="System.String" />
            <Content Name="Signature">
              <OpenSpan.Automation.MethodSignature>
                <ReturnType Value="System.String" />
                <Content Name="ParameterPrototype">
                  <Items>
                    <OpenSpan.Automation.ParameterPrototype>
                      <CanRead Value="False" />
                      <CanWrite Value="True" />
                      <DefaultSet Value="False" />
                      <ParamName Value="number" />
                      <Position Value="0" />
                      <TypeName Value="System.Int32" />
                    </OpenSpan.Automation.ParameterPrototype>
                  </Items>
                </Content>
              </OpenSpan.Automation.MethodSignature>
            </Content>
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableMethod>
    <OpenSpan.Design.TypeProxy Name="dataTableProxy1" Id="TypeProxy-8D5B6618A9CAB58">
      <AliasName Value="" />
      <ProxiedTypeName Value="System.Data.DataTable, System.Data" />
      <UseAlias Value="False" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Instance" canRead="True" canWrite="True" type="System.Data.DataTable" aliasName="Instance" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Design.TypeProxy>
    <OpenSpan.Automation.ConnectableTypeProxy Name="connectableTypeProxy5" Id="ConnectableTypeProxy-8D5B6618A9FDFBD">
      <ComponentName Value="dataTableProxy1" />
      <DisplayName Value="Proxy" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Design.TypeProxy" />
      <InstanceUniqueId Value="Automator-8D540A8F1FCACEA\TypeProxy-8D5B6618A9CAB58" />
      <MemberDetails Value="" />
      <ProxiedTypeName Value="System.Data.DataTable" />
    </OpenSpan.Automation.ConnectableTypeProxy>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties9" Id="ConnectableProperties-8D5B661A969AC8E">
      <ComponentName Value="dataTableProxy1" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="System.Data.DataTable" />
      <InstanceUniqueId Value="Automator-8D540A8F1FCACEA\TypeProxy-8D5B6618A9CAB58" />
      <MemberDetails Value=".Rows Property" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="Rows" />
            <MemberType Value="Property" />
            <TypeAssemblyName Value="System.Data" />
            <TypeName Value="System.Data.DataRowCollection" />
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableProperties>
    <OpenSpan.Controls.ListLoop Name="listLoop1" Id="ListLoop-8D5B661B2C479AB">
      <ComponentName Value="listLoop1" />
      <DisplayName Value="StartLoop" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Controls.ListLoop" />
      <InstanceUniqueId Value="Automator-8D540A8F1FCACEA\ListLoop-8D5B661B2C479AB" />
      <MemberDetails Value="" />
      <Scope Value="Local" Extended="True" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="TypedItem" canRead="True" canWrite="False" type="System.Object" aliasName="TypedItem" shouldSerialize="False" visibility="DefaultOn" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Controls.ListLoop>
    <OpenSpan.Automation.ConnectableMethod Name="connectableMethod7" Id="ConnectableMethod-8D5B6625BF166D4">
      <ComponentName Value="Demo_P_Cast_To_DataRow" />
      <DisplayName Value="Execute" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Automation.Automator" />
      <InstanceUniqueId Value="Automator-8D5B661C6B58228" />
      <MemberDetails Value=".Execute() Method" />
      <ParamsLength Value="0" />
      <SerializedParamsDefaultValues Value="" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Result" canRead="True" canWrite="False" type="System.Data.DataRow" aliasName="Result" shouldSerialize="False" visibility="DefaultOn" source="" blockTypeName="" />
        </Items>
      </Content>
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="_EntryPointExecute" />
            <MemberType Value="Method" />
            <TypeAssemblyName Value="System.Data" />
            <TypeName Value="System.Data.DataRow" />
            <Content Name="Signature">
              <OpenSpan.Automation.MethodSignature>
                <ReturnType Value="System.Data.DataRow" />
                <Content Name="ParameterPrototype">
                  <Items>
                    <OpenSpan.Automation.ParameterPrototype>
                      <CanRead Value="False" />
                      <CanWrite Value="True" />
                      <DefaultSet Value="False" />
                      <ParamName Value="param1" />
                      <Position Value="0" />
                      <TypeAssemblyName Value="System.Data" />
                      <TypeName Value="System.Data.DataRow" />
                    </OpenSpan.Automation.ParameterPrototype>
                  </Items>
                </Content>
              </OpenSpan.Automation.MethodSignature>
            </Content>
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableMethod>
    <OpenSpan.Design.TypeProxy Name="dataRowProxy1" Id="TypeProxy-8D5B662672FCD62">
      <AliasName Value="" />
      <ProxiedTypeName Value="System.Data.DataRow, System.Data" />
      <UseAlias Value="False" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Instance" canRead="True" canWrite="True" type="System.Data.DataRow" aliasName="Instance" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Design.TypeProxy>
    <OpenSpan.Automation.ConnectableTypeProxy Name="connectableTypeProxy6" Id="ConnectableTypeProxy-8D5B6626732DAB6">
      <ComponentName Value="dataRowProxy1" />
      <DisplayName Value="Proxy" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Design.TypeProxy" />
      <InstanceUniqueId Value="Automator-8D540A8F1FCACEA\TypeProxy-8D5B662672FCD62" />
      <MemberDetails Value="" />
      <ProxiedTypeName Value="System.Data.DataRow" />
    </OpenSpan.Automation.ConnectableTypeProxy>
    <OpenSpan.Automation.ConnectableMethod Name="connectableMethod8" Id="ConnectableMethod-8D5B662825A1495">
      <ComponentName Value="dataRowProxy1" />
      <DisplayName Value="SetItem" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="System.Data.DataRow" />
      <InstanceUniqueId Value="Automator-8D540A8F1FCACEA\TypeProxy-8D5B662672FCD62" />
      <MemberDetails Value=".SetItem() Method" />
      <ParamsLength Value="0" />
      <SerializedParamsDefaultValues Value="" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="SetItem" />
            <MemberType Value="Method" />
            <TypeName Value="System.Void" />
            <Content Name="Signature">
              <OpenSpan.Automation.MethodSignature>
                <ReturnType Value="System.Void" />
                <Content Name="ParameterPrototype">
                  <Items>
                    <OpenSpan.Automation.ParameterPrototype>
                      <CanRead Value="False" />
                      <CanWrite Value="True" />
                      <DefaultSet Value="True" />
                      <DefaultValue Value="6" />
                      <ParamName Value="columnIndex" />
                      <Position Value="0" />
                      <TypeName Value="System.Int32" />
                    </OpenSpan.Automation.ParameterPrototype>
                    <OpenSpan.Automation.ParameterPrototype>
                      <CanRead Value="False" />
                      <CanWrite Value="True" />
                      <DefaultSet Value="False" />
                      <ParamName Value="value" />
                      <Position Value="1" />
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
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties10" Id="ConnectableProperties-8D5B6629D9F3778">
      <ComponentName Value="startCell" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Controls.StringVariable" />
      <InstanceUniqueId Value="Automator-8D540A8F1FCACEA\StringVariable-8D5B65E8FC2E8BD" />
      <MemberDetails Value=".Value Property" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="Value" />
            <MemberType Value="Property" />
            <TypeName Value="System.String" />
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableProperties>
    <OpenSpan.Automation.ConnectableMethod Name="connectableMethod9" Id="ConnectableMethod-8D5B6630986ACFF">
      <ComponentName Value="microsoftExcel1" />
      <DisplayName Value="ImportData" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Office.MicrosoftExcel" />
      <InstanceUniqueId Value="GlobalContainer-8D53AA1A011B17D\MicrosoftExcel-8D53AA1A49B9D1B" />
      <MemberDetails Value=".ImportData() Method" />
      <ParamsLength Value="0" />
      <SerializedParamsDefaultValues Value="" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="ImportData" />
            <MemberType Value="Method" />
            <TypeName Value="System.Void" />
            <Content Name="Signature">
              <OpenSpan.Automation.MethodSignature>
                <ReturnType Value="System.Void" />
                <Content Name="ParameterPrototype">
                  <Items>
                    <OpenSpan.Automation.ParameterPrototype>
                      <CanRead Value="False" />
                      <CanWrite Value="True" />
                      <DefaultSet Value="False" />
                      <ParamName Value="dataTable" />
                      <Position Value="0" />
                      <TypeAssemblyName Value="System.Data" />
                      <TypeName Value="System.Data.DataTable" />
                    </OpenSpan.Automation.ParameterPrototype>
                    <OpenSpan.Automation.ParameterPrototype>
                      <CanRead Value="False" />
                      <CanWrite Value="True" />
                      <DefaultSet Value="True" />
                      <DefaultValue Value="True" />
                      <ParamName Value="createHeader" />
                      <Position Value="1" />
                      <TypeName Value="System.Boolean" />
                    </OpenSpan.Automation.ParameterPrototype>
                    <OpenSpan.Automation.ParameterPrototype>
                      <CanRead Value="False" />
                      <CanWrite Value="True" />
                      <DefaultSet Value="True" />
                      <DefaultValue Value="0" />
                      <ParamName Value="rowStart" />
                      <Position Value="2" />
                      <TypeName Value="System.Int32" />
                    </OpenSpan.Automation.ParameterPrototype>
                    <OpenSpan.Automation.ParameterPrototype>
                      <CanRead Value="False" />
                      <CanWrite Value="True" />
                      <DefaultSet Value="True" />
                      <DefaultValue Value="0" />
                      <ParamName Value="columnStart" />
                      <Position Value="3" />
                      <TypeName Value="System.Int32" />
                    </OpenSpan.Automation.ParameterPrototype>
                    <OpenSpan.Automation.ParameterPrototype>
                      <CanRead Value="False" />
                      <CanWrite Value="True" />
                      <DefaultSet Value="True" />
                      <DefaultValue Value="True" />
                      <ParamName Value="uiUpdates" />
                      <Position Value="4" />
                      <TypeName Value="System.Boolean" />
                    </OpenSpan.Automation.ParameterPrototype>
                  </Items>
                </Content>
              </OpenSpan.Automation.MethodSignature>
            </Content>
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableMethod>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties11" Id="ConnectableProperties-8D5B66319EEECC3">
      <ComponentName Value="dataTableProxy1" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="System.Data.DataTable" />
      <InstanceUniqueId Value="Automator-8D540A8F1FCACEA\TypeProxy-8D5B6618A9CAB58" />
      <MemberDetails Value=".This Property" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="This" />
            <MemberType Value="Property" />
            <TypeAssemblyName Value="System.Data" />
            <TypeName Value="System.Data.DataTable" />
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableProperties>
    <OpenSpan.Automation.ConnectableMethod Name="connectableMethod10" Id="ConnectableMethod-8D5B6638DBB9CD7">
      <ComponentName Value="microsoftExcel1" />
      <DisplayName Value="Close" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Office.MicrosoftExcel" />
      <InstanceUniqueId Value="GlobalContainer-8D53AA1A011B17D\MicrosoftExcel-8D53AA1A49B9D1B" />
      <MemberDetails Value=".Close() Method" />
      <ParamsLength Value="0" />
      <SerializedParamsDefaultValues Value="" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="Close" />
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
  </Component>
</OpenSpanDesignDocument>