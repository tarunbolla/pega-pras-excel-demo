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
      <Assembly Value="OpenSpan.Script, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" />
    </AssemblyReferences>
    <DynamicAssemblyReferences>
      <Assembly Value="BooleanExpression-8D53AED08C2FD90" Type="Dynamic.BooleanExpression_8D53AED08C2FD90.Expression" />
    </DynamicAssemblyReferences>
    <FileReferences />
    <BuildReferences />
  </ComponentInfo>
  <Component Version="1.0">
    <OpenSpan.Automation.Automator Name="Demo_P_GetSheetByName" Id="Automator-8D53AEC9D8E3A30">
      <AutomationDocument>
        <Name Value="" />
        <Size Value="5021, 5004" />
        <Objects>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D53AEC9D8E3A30\ConnectableProperties-8D53AECD88B3D90" />
            <PartID Value="2" />
            <Left Value="240" />
            <Top Value="160" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="microsoftExcel1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Proxy" />
            <ConnectableUniqueId Value="Automator-8D53AEC9D8E3A30\ConnectableTypeProxy-8D53AECDC14D980" />
            <PartID Value="3" />
            <Left Value="280" />
            <Top Value="260" />
            <Collapsed Value="False" />
            <WillExecute Value="False" />
            <InstanceName Value="_WorkbookProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D53AEC9D8E3A30\ConnectableProperties-8D53AECECB91A80" />
            <PartID Value="6" />
            <Left Value="460" />
            <Top Value="160" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="_WorkbookProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="StartLoop" />
            <ConnectableUniqueId Value="Automator-8D53AEC9D8E3A30\ListLoop-8D53AECF99EA5D0" />
            <PartID Value="8" />
            <Left Value="680" />
            <Top Value="160" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="listLoop1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock type="OpenSpan.Automation.Design.ConnectionBlocks.EntryPointExecuteBlock">
            <DisplayName Value="Execute" />
            <ConnectableUniqueId Value="Automator-8D53AEC9D8E3A30\ConnectableMethod-8D53AED24A55B70" />
            <PartID Value="16" />
            <Left Value="860" />
            <Top Value="160" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="Demo_P_Cast_To_WorkSheet" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Proxy" />
            <ConnectableUniqueId Value="Automator-8D53AEC9D8E3A30\ConnectableTypeProxy-8D53AED2E096120" />
            <PartID Value="19" />
            <Left Value="860" />
            <Top Value="280" />
            <Collapsed Value="False" />
            <WillExecute Value="False" />
            <InstanceName Value="_WorksheetProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D53AEC9D8E3A30\ConnectableProperties-8D53AEDE2972690" />
            <PartID Value="35" />
            <Left Value="940" />
            <Top Value="60" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="_WorksheetProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <ConnectableUniqueId Value="Automator-8D53AEC9D8E3A30\LabelHost-8D53AEE8B74E9E6" />
            <Left Value="200" />
            <Top Value="20" />
            <PartID Value="38" />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="GoToLabel" />
            <ConnectableUniqueId Value="Automator-8D53AEC9D8E3A30\JumpHost-8D53AEE954D19F8" />
            <PartID Value="40" />
            <Left Value="694" />
            <Top Value="313" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="Jump To" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Equals" />
            <ConnectableUniqueId Value="Automator-8D53AEC9D8E3A30\ConnectableMethod-8D53AF0D8CE9138" />
            <PartID Value="47" />
            <Left Value="1140" />
            <Top Value="160" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="System.String" />
            <Fittings>
              <Result Collapsed="False" ActualText="Result" />
            </Fittings>
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D53AEC9D8E3A30\ConnectableProperties-8D53AED74530050" />
            <PartID Value="26" />
            <Left Value="1160" />
            <Top Value="60" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="sheetName" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <ConnectableUniqueId Value="Automator-8D53AEC9D8E3A30\EntryPoint-8D53AECA06DB370" />
            <Left Value="63" />
            <Top Value="162" />
            <PartID Value="1" />
          </ConnectionBlock>
          <ConnectionBlock type="OpenSpan.Automation.Design.ConnectionBlocks.MultiExitPointBlock">
            <ConnectableUniqueId Value="Automator-8D53AEC9D8E3A30\ExitPoint-8D53AEE897AB120" />
            <Left Value="383" />
            <Top Value="22" />
            <PartID Value="37" />
            <Title Value="Exit" />
            <EventName Value="" />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="GoToLabel" />
            <ConnectableUniqueId Value="Automator-8D53AEC9D8E3A30\JumpHost-8D53B1565E96DB0" />
            <PartID Value="54" />
            <Left Value="1060" />
            <Top Value="360" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="Jump To" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D53AEC9D8E3A30\ConnectableProperties-8D53B1583340E20" />
            <PartID Value="57" />
            <Left Value="860" />
            <Top Value="420" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="_WorksheetProxy1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Break" />
            <ConnectableUniqueId Value="Automator-8D53AEC9D8E3A30\ConnectableMethod-8D53B16858F1FB0" />
            <PartID Value="63" />
            <Left Value="1340" />
            <Top Value="160" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="listLoop1" />
            <OverriddenIds />
          </ConnectionBlock>
        </Objects>
        <Links>
          <Link PartID="4" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="2" PortName="ExcelWorkbook" PortType="Property" ConnectableId="Automator-8D53AEC9D8E3A30\ConnectableProperties-8D53AECD88B3D90" MemberComponentId="GlobalContainer-8D53AA1A011B17D\MicrosoftExcel-8D53AA1A49B9D1B" />
            <To PartID="3" PortName="Instance" PortType="Property" ConnectableId="Automator-8D53AEC9D8E3A30\ConnectableTypeProxy-8D53AECDC14D980" MemberComponentId="Automator-8D53AEC9D8E3A30\TypeProxy-8D53AECDC0FA960" />
            <LinkPoints>
              <Point value="386, 206" />
              <Point value="396, 206" />
              <Point value="396, 206" />
              <Point value="396, 228" />
              <Point value="276, 228" />
              <Point value="276, 305" />
              <Point value="275, 305" />
              <Point value="285, 305" />
            </LinkPoints>
          </Link>
          <Link PartID="7" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="2" PortName="Complete" PortType="Event" ConnectableId="Automator-8D53AEC9D8E3A30\ConnectableProperties-8D53AECD88B3D90" MemberComponentId="Automator-8D53AEC9D8E3A30\ConnectableProperties-8D53AECD88B3D90" />
            <To PartID="6" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AEC9D8E3A30\ConnectableProperties-8D53AECECB91A80" MemberComponentId="Automator-8D53AEC9D8E3A30\ConnectableProperties-8D53AECECB91A80" />
            <LinkPoints>
              <Point value="386, 189" />
              <Point value="396, 189" />
              <Point value="396, 189" />
              <Point value="396, 189" />
              <Point value="455, 189" />
              <Point value="465, 189" />
            </LinkPoints>
          </Link>
          <Link PartID="9" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="6" PortName="Complete" PortType="Event" ConnectableId="Automator-8D53AEC9D8E3A30\ConnectableProperties-8D53AECECB91A80" MemberComponentId="Automator-8D53AEC9D8E3A30\ConnectableProperties-8D53AECECB91A80" />
            <To PartID="8" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AEC9D8E3A30\ListLoop-8D53AECF99EA5D0" MemberComponentId="Automator-8D53AEC9D8E3A30\ListLoop-8D53AECF99EA5D0" />
            <LinkPoints>
              <Point value="621, 189" />
              <Point value="631, 189" />
              <Point value="631, 189" />
              <Point value="631, 189" />
              <Point value="675, 189" />
              <Point value="685, 189" />
            </LinkPoints>
          </Link>
          <Link PartID="10" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="6" PortName="Worksheets" PortType="Property" ConnectableId="Automator-8D53AEC9D8E3A30\ConnectableProperties-8D53AECECB91A80" MemberComponentId="Automator-8D53AEC9D8E3A30\TypeProxy-8D53AECDC0FA960" />
            <To PartID="8" PortName="List" PortType="Property" ConnectableId="Automator-8D53AEC9D8E3A30\ListLoop-8D53AECF99EA5D0" MemberComponentId="Automator-8D53AEC9D8E3A30\ListLoop-8D53AECF99EA5D0" />
            <LinkPoints>
              <Point value="621, 206" />
              <Point value="631, 206" />
              <Point value="631, 206" />
              <Point value="631, 206" />
              <Point value="675, 206" />
              <Point value="685, 206" />
            </LinkPoints>
          </Link>
          <Link PartID="20" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="16" PortName="Result" PortType="Property" ConnectableId="Automator-8D53AEC9D8E3A30\ConnectableMethod-8D53AED24A55B70" MemberComponentId="Automator-8D53AEC9D8E3A30\ConnectableMethod-8D53AED24A55B70" />
            <To PartID="19" PortName="Instance" PortType="Property" ConnectableId="Automator-8D53AEC9D8E3A30\ConnectableTypeProxy-8D53AED2E096120" MemberComponentId="Automator-8D53AEC9D8E3A30\TypeProxy-8D53AED2E059090" />
            <LinkPoints>
              <Point value="1097, 240" />
              <Point value="1107, 240" />
              <Point value="1108, 240" />
              <Point value="1108, 260" />
              <Point value="852, 260" />
              <Point value="852, 325" />
              <Point value="855, 325" />
              <Point value="865, 325" />
            </LinkPoints>
          </Link>
          <Link PartID="22" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="8" PortName="ItemYielded" PortType="Event" ConnectableId="Automator-8D53AEC9D8E3A30\ListLoop-8D53AECF99EA5D0" MemberComponentId="Automator-8D53AEC9D8E3A30\ListLoop-8D53AECF99EA5D0" />
            <To PartID="16" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AEC9D8E3A30\ConnectableMethod-8D53AED24A55B70" MemberComponentId="Automator-8D53AEC9D8E3A30\ConnectableMethod-8D53AED24A55B70" />
            <LinkPoints>
              <Point value="799, 223" />
              <Point value="809, 223" />
              <Point value="812, 223" />
              <Point value="812, 189" />
              <Point value="855, 189" />
              <Point value="865, 189" />
            </LinkPoints>
          </Link>
          <Link PartID="23" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="8" PortName="TypedItem" PortType="Property" ConnectableId="Automator-8D53AEC9D8E3A30\ListLoop-8D53AECF99EA5D0" MemberComponentId="Automator-8D53AEC9D8E3A30\ListLoop-8D53AECF99EA5D0" />
            <To PartID="16" PortName="param1" PortType="Property" ConnectableId="Automator-8D53AEC9D8E3A30\ConnectableMethod-8D53AED24A55B70" MemberComponentId="Automator-8D53AEC9D8E3A30\ConnectableMethod-8D53AED24A55B70" />
            <LinkPoints>
              <Point value="799, 240" />
              <Point value="809, 240" />
              <Point value="812, 240" />
              <Point value="812, 223" />
              <Point value="855, 223" />
              <Point value="865, 223" />
            </LinkPoints>
          </Link>
          <Link PartID="41" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="8" PortName="Completed" PortType="Event" ConnectableId="Automator-8D53AEC9D8E3A30\ListLoop-8D53AECF99EA5D0" MemberComponentId="Automator-8D53AEC9D8E3A30\ListLoop-8D53AECF99EA5D0" />
            <To PartID="40" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AEC9D8E3A30\JumpHost-8D53AEE954D19F8" MemberComponentId="Automator-8D53AEC9D8E3A30\JumpHost-8D53AEE954D19F8" />
            <LinkPoints>
              <Point value="799, 274" />
              <Point value="809, 274" />
              <Point value="812, 274" />
              <Point value="812, 292" />
              <Point value="684, 292" />
              <Point value="684, 330" />
              <Point value="687, 330" />
              <Point value="697, 330" />
            </LinkPoints>
          </Link>
          <Link PartID="49" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="16" PortName="Complete" PortType="Event" ConnectableId="Automator-8D53AEC9D8E3A30\ConnectableMethod-8D53AED24A55B70" MemberComponentId="Automator-8D53AA3CD4DFDD0\ExitPoint-8D53AA3DB396060" />
            <To PartID="47" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AEC9D8E3A30\ConnectableMethod-8D53AF0D8CE9138" MemberComponentId="Automator-8D53AEC9D8E3A30\ConnectableMethod-8D53AF0D8CE9138" />
            <LinkPoints>
              <Point value="1097, 206" />
              <Point value="1107, 206" />
              <Point value="1108, 206" />
              <Point value="1108, 189" />
              <Point value="1135, 189" />
              <Point value="1145, 189" />
            </LinkPoints>
          </Link>
          <Link PartID="50" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="35" PortName="Name" PortType="Property" ConnectableId="Automator-8D53AEC9D8E3A30\ConnectableProperties-8D53AEDE2972690" MemberComponentId="Automator-8D53AEC9D8E3A30\TypeProxy-8D53AED2E059090" />
            <To PartID="47" PortName="b" PortType="Property" ConnectableId="Automator-8D53AEC9D8E3A30\ConnectableMethod-8D53AF0D8CE9138" MemberComponentId="Automator-8D53AEC9D8E3A30\ConnectableMethod-8D53AF0D8CE9138" />
            <LinkPoints>
              <Point value="1105, 106" />
              <Point value="1115, 106" />
              <Point value="1116, 106" />
              <Point value="1116, 223" />
              <Point value="1135, 223" />
              <Point value="1145, 223" />
            </LinkPoints>
          </Link>
          <Link PartID="5" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="1" PortName="Complete" PortType="Event" ConnectableId="Automator-8D53AEC9D8E3A30\EntryPoint-8D53AECA06DB370" MemberComponentId="Automator-8D53AEC9D8E3A30\EntryPoint-8D53AECA06DB370" />
            <To PartID="2" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AEC9D8E3A30\ConnectableProperties-8D53AECD88B3D90" MemberComponentId="Automator-8D53AEC9D8E3A30\ConnectableProperties-8D53AECD88B3D90" />
            <LinkPoints>
              <Point value="177, 178" />
              <Point value="187, 178" />
              <Point value="188, 178" />
              <Point value="188, 189" />
              <Point value="235, 189" />
              <Point value="245, 189" />
            </LinkPoints>
          </Link>
          <Link PartID="51" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="26" PortName="This" PortType="Property" ConnectableId="Automator-8D53AEC9D8E3A30\ConnectableProperties-8D53AED74530050" MemberComponentId="Automator-8D53AEC9D8E3A30\HiddenTypeProxy-8D53AECAA66CD30" />
            <To PartID="47" PortName="a" PortType="Property" ConnectableId="Automator-8D53AEC9D8E3A30\ConnectableMethod-8D53AF0D8CE9138" MemberComponentId="Automator-8D53AEC9D8E3A30\ConnectableMethod-8D53AF0D8CE9138" />
            <LinkPoints>
              <Point value="1279, 106" />
              <Point value="1289, 106" />
              <Point value="1292, 106" />
              <Point value="1292, 124" />
              <Point value="1132, 124" />
              <Point value="1132, 206" />
              <Point value="1135, 206" />
              <Point value="1145, 206" />
            </LinkPoints>
          </Link>
          <Link PartID="39" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="38" PortName="Complete" PortType="Event" ConnectableId="Automator-8D53AEC9D8E3A30\LabelHost-8D53AEE8B74E9E6" MemberComponentId="Automator-8D53AEC9D8E3A30\LabelHost-8D53AEE8B74E9E6" />
            <To PartID="37" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AEC9D8E3A30\ExitPoint-8D53AEE897AB120" MemberComponentId="Automator-8D53AEC9D8E3A30\ExitPoint-8D53AEE897AB120" />
            <LinkPoints>
              <Point value="323, 38" />
              <Point value="333, 38" />
              <Point value="333, 38" />
              <Point value="333, 40" />
              <Point value="375, 40" />
              <Point value="385, 40" />
            </LinkPoints>
          </Link>
          <Link PartID="61" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="57" PortName="This" PortType="Property" ConnectableId="Automator-8D53AEC9D8E3A30\ConnectableProperties-8D53B1583340E20" MemberComponentId="Automator-8D53AEC9D8E3A30\TypeProxy-8D53AED2E059090" />
            <To PartID="54" PortName="_param1" PortType="Property" ConnectableId="Automator-8D53AEC9D8E3A30\JumpHost-8D53B1565E96DB0" MemberComponentId="Automator-8D53AEC9D8E3A30\JumpHost-8D53B1565E96DB0" />
            <LinkPoints>
              <Point value="1025, 466" />
              <Point value="1035, 466" />
              <Point value="1036, 466" />
              <Point value="1036, 406" />
              <Point value="1055, 406" />
              <Point value="1065, 406" />
            </LinkPoints>
          </Link>
          <Link PartID="62" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="8" PortName="LoopBroken" PortType="Event" ConnectableId="Automator-8D53AEC9D8E3A30\ListLoop-8D53AECF99EA5D0" MemberComponentId="Automator-8D53AEC9D8E3A30\ListLoop-8D53AECF99EA5D0" />
            <To PartID="54" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AEC9D8E3A30\JumpHost-8D53B1565E96DB0" MemberComponentId="Automator-8D53AEC9D8E3A30\JumpHost-8D53B1565E96DB0" />
            <LinkPoints>
              <Point value="799, 257" />
              <Point value="809, 257" />
              <Point value="812, 257" />
              <Point value="812, 377" />
              <Point value="1053, 377" />
              <Point value="1063, 377" />
            </LinkPoints>
          </Link>
          <DecisionEventLink PartID="64" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="47" ParentMemberName="Result" DecisionValue="True" ConnectableId="Automator-8D53AEC9D8E3A30\ConnectableMethod-8D53AF0D8CE9138" />
            <To PartID="63" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AEC9D8E3A30\ConnectableMethod-8D53B16858F1FB0" MemberComponentId="Automator-8D53AEC9D8E3A30\ConnectableMethod-8D53B16858F1FB0" />
            <LinkPoints>
              <Point value="1274, 254" />
              <Point value="1284, 254" />
              <Point value="1310, 254" />
              <Point value="1310, 189" />
              <Point value="1335, 189" />
              <Point value="1345, 189" />
            </LinkPoints>
          </DecisionEventLink>
          <Link PartID="65" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="38" PortName="_param1" PortType="Property" ConnectableId="Automator-8D53AEC9D8E3A30\LabelHost-8D53AEE8B74E9E6" MemberComponentId="EMPTY" />
            <To PartID="37" PortName="Result" PortType="Property" ConnectableId="Automator-8D53AEC9D8E3A30\ExitPoint-8D53AEE897AB120" MemberComponentId="EMPTY" />
            <LinkPoints>
              <Point value="323, 65" />
              <Point value="333, 65" />
              <Point value="354, 65" />
              <Point value="354, 67" />
              <Point value="375, 67" />
              <Point value="385, 67" />
            </LinkPoints>
          </Link>
        </Links>
        <Comments />
        <SubGraphs />
      </AutomationDocument>
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicMethodInfo dynamicType="Method" name="_EntryPointExecute" aliasName="Execute" visibility="DefaultOn" source="" blockTypeName="OpenSpan.Automation.Design.ConnectionBlocks.EntryPointExecuteBlock" returnTypeAssembly="Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" returnType="Microsoft.Office.Interop.Excel._Worksheet">
            <param name="_param1" aliasName="sheetName" paramType="System.String" isIn="True" isOut="False" position="0" />
          </OpenSpan.DynamicMembers.DynamicMethodInfo>
        </Items>
      </Content>
    </OpenSpan.Automation.Automator>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties1" Id="ConnectableProperties-8D53AECD88B3D90">
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
    <OpenSpan.Design.TypeProxy Name="_WorkbookProxy1" Id="TypeProxy-8D53AECDC0FA960">
      <AliasName Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel._Workbook, Microsoft.Office.Interop.Excel" />
      <UseAlias Value="False" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Instance" canRead="True" canWrite="True" typeAssembly="Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" type="Microsoft.Office.Interop.Excel._Workbook" aliasName="Instance" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Design.TypeProxy>
    <OpenSpan.Automation.ConnectableTypeProxy Name="connectableTypeProxy1" Id="ConnectableTypeProxy-8D53AECDC14D980">
      <ComponentName Value="_WorkbookProxy1" />
      <DisplayName Value="Proxy" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel._Workbook" />
      <InstanceUniqueId Value="Automator-8D53AEC9D8E3A30\TypeProxy-8D53AECDC0FA960" />
      <MemberDetails Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel._Workbook" />
    </OpenSpan.Automation.ConnectableTypeProxy>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties2" Id="ConnectableProperties-8D53AECECB91A80">
      <ComponentName Value="_WorkbookProxy1" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel._Workbook" />
      <InstanceUniqueId Value="Automator-8D53AEC9D8E3A30\TypeProxy-8D53AECDC0FA960" />
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
    <OpenSpan.Controls.ListLoop Name="listLoop1" Id="ListLoop-8D53AECF99EA5D0">
      <ComponentName Value="listLoop1" />
      <DisplayName Value="StartLoop" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Controls.ListLoop" />
      <InstanceUniqueId Value="Automator-8D53AEC9D8E3A30\ListLoop-8D53AECF99EA5D0" />
      <MemberDetails Value="" />
      <Scope Value="Local" Extended="True" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="TypedItem" canRead="True" canWrite="False" type="System.Object" aliasName="TypedItem" shouldSerialize="False" visibility="DefaultOn" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Controls.ListLoop>
    <OpenSpan.Script.Expression.BooleanExpression Name="booleanExpression2" Id="BooleanExpression-8D53AED08C2FD90">
      <Expression Value="a==b" />
      <Scope Value="Local" Extended="True" />
      <Valid Value="True" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicMethodInfo dynamicType="Method" name="Evaluate" aliasName="Evaluate" visibility="DefaultOn" source="" blockTypeName="" returnType="System.Boolean">
            <param name="a" aliasName="a" paramType="System.Double" isIn="True" isOut="False" position="0" />
            <param name="b" aliasName="b" paramType="System.Double" isIn="True" isOut="False" position="1" />
          </OpenSpan.DynamicMembers.DynamicMethodInfo>
        </Items>
      </Content>
      <Content Name="Identifiers">
        <Items>
          <OpenSpan.Script.Expression.ExpressionIdentifier>
            <DataType Value="Double" />
            <ID Value="a" />
          </OpenSpan.Script.Expression.ExpressionIdentifier>
          <OpenSpan.Script.Expression.ExpressionIdentifier>
            <DataType Value="Double" />
            <ID Value="b" />
          </OpenSpan.Script.Expression.ExpressionIdentifier>
        </Items>
      </Content>
    </OpenSpan.Script.Expression.BooleanExpression>
    <OpenSpan.Automation.ConnectableMethod Name="connectableMethod1" Id="ConnectableMethod-8D53AED24A55B70">
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
    <OpenSpan.Design.TypeProxy Name="_WorksheetProxy1" Id="TypeProxy-8D53AED2E059090">
      <AliasName Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel._Worksheet, Microsoft.Office.Interop.Excel" />
      <UseAlias Value="False" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Instance" canRead="True" canWrite="True" typeAssembly="Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" type="Microsoft.Office.Interop.Excel._Worksheet" aliasName="Instance" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Design.TypeProxy>
    <OpenSpan.Automation.ConnectableTypeProxy Name="connectableTypeProxy3" Id="ConnectableTypeProxy-8D53AED2E096120">
      <ComponentName Value="_WorksheetProxy1" />
      <DisplayName Value="Proxy" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel._Worksheet" />
      <InstanceUniqueId Value="Automator-8D53AEC9D8E3A30\TypeProxy-8D53AED2E059090" />
      <MemberDetails Value="" />
      <ProxiedTypeName Value="Microsoft.Office.Interop.Excel._Worksheet" />
    </OpenSpan.Automation.ConnectableTypeProxy>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties3" Id="ConnectableProperties-8D53AEDE2972690">
      <ComponentName Value="_WorksheetProxy1" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel._Worksheet" />
      <InstanceUniqueId Value="Automator-8D53AEC9D8E3A30\TypeProxy-8D53AED2E059090" />
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
    <OpenSpan.Automation.LabelHost Name="labelHost1" Id="LabelHost-8D53AEE8B74E9E6">
      <ComponentName Value="OpenSpan.Automation.EntryPoint" />
      <DisplayName Value="Exit" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Automation.EntryPoint" />
      <InstanceUniqueId Value="EMPTY" />
      <LabelName Value="Exit" />
      <MemberDetails Value="" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="_param1" canRead="True" canWrite="True" typeAssembly="Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" type="Microsoft.Office.Interop.Excel._Worksheet" aliasName="Result" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
          <OpenSpan.DynamicMembers.DynamicMethodInfo dynamicType="Method" name="GoToLabel" aliasName="GoToLabel" visibility="AlwaysHidden" source="" blockTypeName="" returnType="System.Void">
            <param name="_param1" aliasName="Result" paramTypeAssembly="Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" paramType="Microsoft.Office.Interop.Excel._Worksheet" isIn="True" isOut="False" position="0" />
          </OpenSpan.DynamicMembers.DynamicMethodInfo>
        </Items>
      </Content>
    </OpenSpan.Automation.LabelHost>
    <OpenSpan.Automation.JumpHost Name="jumpHost1" Id="JumpHost-8D53AEE954D19F8">
      <ComponentName Value="labelHost1" />
      <DisplayName Value="GoToLabel" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Automation.LabelHost" />
      <InstanceUniqueId Value="Automator-8D53AEC9D8E3A30\LabelHost-8D53AEE8B74E9E6" />
      <MemberDetails Value=" - Exit" />
      <ParamsLength Value="0" />
      <SerializedParamsDefaultValues Value="" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="GoToLabel" />
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
                      <ParamName Value="_param1" />
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
    </OpenSpan.Automation.JumpHost>
    <OpenSpan.Controls.MessageDialog Name="messageDialog1" Id="MessageDialog-8D53AEEA8960070">
      <Caption Value="Information" />
      <Scope Value="Local" Extended="True" />
    </OpenSpan.Controls.MessageDialog>
    <OpenSpan.Controls.MessageDialog Name="messageDialog2" Id="MessageDialog-8D53AEEAA56DA3A">
      <Caption Value="Information" />
      <Scope Value="Local" Extended="True" />
    </OpenSpan.Controls.MessageDialog>
    <OpenSpan.Automation.ConnectableMethod Name="connectableMethod5" Id="ConnectableMethod-8D53AF0D8CE9138">
      <ComponentName Value="System.String" />
      <DisplayName Value="Equals" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="System.String" />
      <InstanceUniqueId Value="EMPTY" />
      <MemberDetails Value=".Equals() Method" />
      <ParamsLength Value="0" />
      <SerializedParamsDefaultValues Value="" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Result" canRead="True" canWrite="False" type="System.Boolean" aliasName="Result" shouldSerialize="False" visibility="DefaultOn" source="" blockTypeName="" />
        </Items>
      </Content>
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="Equals" />
            <MemberType Value="Method" />
            <TypeName Value="System.Boolean" />
            <Content Name="Signature">
              <OpenSpan.Automation.MethodSignature>
                <ReturnType Value="System.Boolean" />
                <Content Name="ParameterPrototype">
                  <Items>
                    <OpenSpan.Automation.ParameterPrototype>
                      <CanRead Value="False" />
                      <CanWrite Value="True" />
                      <DefaultSet Value="False" />
                      <ParamName Value="a" />
                      <Position Value="0" />
                      <TypeName Value="System.String" />
                    </OpenSpan.Automation.ParameterPrototype>
                    <OpenSpan.Automation.ParameterPrototype>
                      <CanRead Value="False" />
                      <CanWrite Value="True" />
                      <DefaultSet Value="False" />
                      <ParamName Value="b" />
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
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties4" Id="ConnectableProperties-8D53AED74530050">
      <ComponentName Value="sheetName" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="System.String" />
      <InstanceUniqueId Value="Automator-8D53AEC9D8E3A30\HiddenTypeProxy-8D53AECAA66CD30" />
      <MemberDetails Value=".This Property" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="This" />
            <MemberType Value="Property" />
            <TypeName Value="System.String" />
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableProperties>
    <OpenSpan.Automation.EntryPoint Name="entryPoint1" Id="EntryPoint-8D53AECA06DB370">
      <AliasName Value="Execute" />
      <ComponentName Value="Execute" />
      <DisplayName Value="Execute" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Automation.EntryPoint" />
      <InstanceUniqueId Value="Automator-8D53AEC9D8E3A30\EntryPoint-8D53AECA06DB370" />
      <MemberDetails Value="" />
      <MethodName Value="_EntryPointExecute" />
      <Removing Value="True" />
      <UseAlias Value="True" />
      <Content Name="Controls">
        <Capacity Value="4" />
        <Items>
          <OpenSpan.Automation.HiddenTypeProxy Name="hiddenTypeProxy2" Id="HiddenTypeProxy-8D53AECAA66CD30">
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
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Result" canRead="True" canWrite="True" typeAssembly="Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" type="Microsoft.Office.Interop.Excel._Worksheet" aliasName="Result" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="_param1" canRead="False" canWrite="True" type="System.String" aliasName="sheetName" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Automation.EntryPoint>
    <OpenSpan.Automation.ExitPoint Name="exitPoint1" Id="ExitPoint-8D53AEE897AB120">
      <ComponentName Value="Execute" />
      <DisplayName Value="Exit" />
      <EntryPoint Value="ComponentReference" Name="entryPoint1" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Automation.EntryPoint" />
      <InstanceUniqueId Value="Automator-8D53AEC9D8E3A30\EntryPoint-8D53AECA06DB370" />
      <MemberDetails Value="" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Result" canRead="True" canWrite="True" typeAssembly="Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" type="Microsoft.Office.Interop.Excel._Worksheet" aliasName="Result" shouldSerialize="False" visibility="DefaultOn" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Automation.ExitPoint>
    <OpenSpan.Automation.JumpHost Name="jumpHost2" Id="JumpHost-8D53B1565E96DB0">
      <ComponentName Value="labelHost1" />
      <DisplayName Value="GoToLabel" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Automation.LabelHost" />
      <InstanceUniqueId Value="Automator-8D53AEC9D8E3A30\LabelHost-8D53AEE8B74E9E6" />
      <MemberDetails Value=" - Exit" />
      <ParamsLength Value="0" />
      <SerializedParamsDefaultValues Value="" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="GoToLabel" />
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
                      <ParamName Value="_param1" />
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
    </OpenSpan.Automation.JumpHost>
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties5" Id="ConnectableProperties-8D53B1583340E20">
      <ComponentName Value="_WorksheetProxy1" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="Microsoft.Office.Interop.Excel._Worksheet" />
      <InstanceUniqueId Value="Automator-8D53AEC9D8E3A30\TypeProxy-8D53AED2E059090" />
      <MemberDetails Value=".This Property" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="This" />
            <MemberType Value="Property" />
            <TypeAssemblyName Value="Microsoft.Office.Interop.Excel" />
            <TypeName Value="Microsoft.Office.Interop.Excel._Worksheet" />
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableProperties>
    <OpenSpan.Automation.ConnectableMethod Name="connectableMethod2" Id="ConnectableMethod-8D53B16858F1FB0">
      <ComponentName Value="listLoop1" />
      <DisplayName Value="Break" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Controls.ListLoop" />
      <InstanceUniqueId Value="Automator-8D53AEC9D8E3A30\ListLoop-8D53AECF99EA5D0" />
      <MemberDetails Value=".Break() Method" />
      <ParamsLength Value="0" />
      <SerializedParamsDefaultValues Value="" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="Break" />
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