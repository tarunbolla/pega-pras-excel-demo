<OpenSpanDesignDocument Version="8.0.1026.0" Serializer="2.0" Culture="en-US">
  <ComponentInfo>
    <Type Value="OpenSpan.Automation.Automator" />
    <Assembly Value="OpenSpan.Automation" />
    <AssemblyReferences>
      <Assembly Value="mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" />
      <Assembly Value="OpenSpan, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" />
      <Assembly Value="OpenSpan.Automation, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" />
      <Assembly Value="OpenSpan.Runtime.Core, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" />
      <Assembly Value="OSComponents.Utilities.TestHarness, Version=8.0.21.0, Culture=neutral, PublicKeyToken=c0cb69f111622a50" />
      <Assembly Value="System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" />
    </AssemblyReferences>
    <DynamicAssemblyReferences />
    <FileReferences />
    <BuildReferences />
  </ComponentInfo>
  <Component Version="1.0">
    <OpenSpan.Automation.Automator Name="Demo_E_AllProjectsStarted" Id="Automator-8D53AEC76F2A6E0">
      <AutomationDocument>
        <Name Value="" />
        <Size Value="5000, 5000" />
        <Objects>
          <ConnectionBlock>
            <DisplayName Value="RuntimeLoader.AllProjectsStarted" />
            <ConnectableUniqueId Value="Automator-8D53AEC76F2A6E0\ConnectableEvent-8D53AEC7DBD15F0" />
            <PartID Value="1" />
            <Left Value="20" />
            <Top Value="40" />
            <Collapsed Value="True" />
            <WillExecute Value="True" />
            <InstanceName Value="OpenSpan.Runtime.RuntimeLoader" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="ShowTestHarness" />
            <ConnectableUniqueId Value="Automator-8D53AEC76F2A6E0\ConnectableMethod-8D53AEC82CFE9F0" />
            <PartID Value="2" />
            <Left Value="380" />
            <Top Value="40" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="testHarness1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D53AEC76F2A6E0\ConnectableProperties-8D53AF0FC514A1B" />
            <PartID Value="4" />
            <Left Value="380" />
            <Top Value="120" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="testHarness1" />
            <Fittings>
              <TopMost Collapsed="False" ActualText="False" />
            </Fittings>
            <OverriddenIds />
          </ConnectionBlock>
        </Objects>
        <Links>
          <Link PartID="3" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="1" PortName="Raised" PortType="Event" ConnectableId="Automator-8D53AEC76F2A6E0\ConnectableEvent-8D53AEC7DBD15F0" MemberComponentId="Automator-8D53AEC76F2A6E0\EMPTY" />
            <To PartID="2" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AEC76F2A6E0\ConnectableMethod-8D53AEC82CFE9F0" MemberComponentId="Automator-8D53AEC76F2A6E0\ConnectableMethod-8D53AEC82CFE9F0" />
            <LinkPoints>
              <Point value="288, 69" />
              <Point value="298, 69" />
              <Point value="336, 69" />
              <Point value="336, 69" />
              <Point value="375, 69" />
              <Point value="385, 69" />
            </LinkPoints>
          </Link>
          <Link PartID="5" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="2" PortName="Complete" PortType="Event" ConnectableId="Automator-8D53AEC76F2A6E0\ConnectableMethod-8D53AEC82CFE9F0" MemberComponentId="Automator-8D53AEC76F2A6E0\ConnectableMethod-8D53AEC82CFE9F0" />
            <To PartID="4" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53AEC76F2A6E0\ConnectableProperties-8D53AF0FC514A1B" MemberComponentId="Automator-8D53AEC76F2A6E0\ConnectableProperties-8D53AF0FC514A1B" />
            <LinkPoints>
              <Point value="537, 69" />
              <Point value="547, 69" />
              <Point value="548, 69" />
              <Point value="548, 84" />
              <Point value="372, 84" />
              <Point value="372, 149" />
              <Point value="375, 149" />
              <Point value="385, 149" />
            </LinkPoints>
          </Link>
        </Links>
        <Comments />
        <SubGraphs />
      </AutomationDocument>
    </OpenSpan.Automation.Automator>
    <OpenSpan.Automation.ConnectableEvent Name="connectableEvent1" Id="ConnectableEvent-8D53AEC7DBD15F0">
      <ComponentName Value="OpenSpan.Runtime.RuntimeLoader" />
      <DisplayName Value="RuntimeLoader.AllProjectsStarted" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Runtime.RuntimeLoader" />
      <InstanceUniqueId Value="EMPTY" />
      <MemberDetails Value=".AllProjectsStarted Event" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="AllProjectsStarted" />
            <MemberType Value="Event" />
            <TypeName Value="System.EventHandler" />
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableEvent>
    <OpenSpan.Automation.ConnectableMethod Name="connectableMethod1" Id="ConnectableMethod-8D53AEC82CFE9F0">
      <ComponentName Value="testHarness1" />
      <DisplayName Value="ShowTestHarness" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OSComponents.Utilities.TestHarness.TestHarness" />
      <InstanceUniqueId Value="GlobalContainer-8D53AA1A011B17D\TestHarness-8D53AEC674CC310" />
      <MemberDetails Value=".ShowTestHarness() Method" />
      <ParamsLength Value="0" />
      <SerializedParamsDefaultValues Value="" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="ShowTestHarness" />
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
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties1" Id="ConnectableProperties-8D53AF0FC514A1B">
      <ComponentName Value="testHarness1" />
      <DefaultValues Value="TopMost=False" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OSComponents.Utilities.TestHarness.TestHarness" />
      <InstanceUniqueId Value="GlobalContainer-8D53AA1A011B17D\TestHarness-8D53AEC674CC310" />
      <MemberDetails Value=".TopMost Property" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="TopMost" />
            <MemberType Value="Property" />
            <TypeName Value="System.Boolean" />
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableProperties>
  </Component>
</OpenSpanDesignDocument>