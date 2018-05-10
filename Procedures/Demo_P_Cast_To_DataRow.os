<OpenSpanDesignDocument Version="8.0.1026.0" Serializer="2.0" Culture="en-US">
  <ComponentInfo>
    <Type Value="OpenSpan.Automation.Automator" />
    <Assembly Value="OpenSpan.Automation" />
    <AssemblyReferences>
      <Assembly Value="mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" />
      <Assembly Value="OpenSpan, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" />
      <Assembly Value="OpenSpan.Automation, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" />
      <Assembly Value="System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" />
    </AssemblyReferences>
    <DynamicAssemblyReferences />
    <FileReferences />
    <BuildReferences />
  </ComponentInfo>
  <Component Version="1.0">
    <OpenSpan.Automation.Automator Name="Demo_P_Cast_To_DataRow" Id="Automator-8D5B661C6B58228">
      <AutomationDocument>
        <Name Value="" />
        <Size Value="5008, 5000" />
        <Objects>
          <ConnectionBlock>
            <ConnectableUniqueId Value="Automator-8D5B661C6B58228\EntryPoint-8D5B661C9B89CB5" />
            <Left Value="60" />
            <Top Value="100" />
            <PartID Value="1" />
          </ConnectionBlock>
          <ConnectionBlock type="OpenSpan.Automation.Design.ConnectionBlocks.MultiExitPointBlock">
            <ConnectableUniqueId Value="Automator-8D5B661C6B58228\ExitPoint-8D5B66247290BA4" />
            <Left Value="240" />
            <Top Value="100" />
            <PartID Value="2" />
            <Title Value="Exit1" />
            <EventName Value="" />
          </ConnectionBlock>
        </Objects>
        <Links>
          <Link PartID="3" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="1" PortName="Complete" PortType="Event" ConnectableId="Automator-8D5B661C6B58228\EntryPoint-8D5B661C9B89CB5" MemberComponentId="Automator-8D5B661C6B58228\EntryPoint-8D5B661C9B89CB5" />
            <To PartID="2" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D5B661C6B58228\ExitPoint-8D5B66247290BA4" MemberComponentId="Automator-8D5B661C6B58228\ExitPoint-8D5B66247290BA4" />
            <LinkPoints>
              <Point value="174, 118" />
              <Point value="184, 118" />
              <Point value="184, 118" />
              <Point value="184, 118" />
              <Point value="232, 118" />
              <Point value="242, 118" />
            </LinkPoints>
          </Link>
          <Link PartID="4" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="1" PortName="param1" PortType="Property" ConnectableId="Automator-8D5B661C6B58228\EntryPoint-8D5B661C9B89CB5" MemberComponentId="EMPTY" />
            <To PartID="2" PortName="Result" PortType="Property" ConnectableId="Automator-8D5B661C6B58228\ExitPoint-8D5B66247290BA4" MemberComponentId="EMPTY" />
            <LinkPoints>
              <Point value="174, 145" />
              <Point value="184, 145" />
              <Point value="208, 145" />
              <Point value="208, 145" />
              <Point value="232, 145" />
              <Point value="242, 145" />
            </LinkPoints>
          </Link>
        </Links>
        <Comments />
        <SubGraphs />
      </AutomationDocument>
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicMethodInfo dynamicType="Method" name="_EntryPointExecute" aliasName="Execute" visibility="DefaultOn" source="" blockTypeName="OpenSpan.Automation.Design.ConnectionBlocks.EntryPointExecuteBlock" returnType="System.Data.DataRow">
            <param name="param1" aliasName="param1" paramType="System.Data.DataRow" isIn="True" isOut="False" position="0" />
          </OpenSpan.DynamicMembers.DynamicMethodInfo>
        </Items>
      </Content>
    </OpenSpan.Automation.Automator>
    <OpenSpan.Automation.EntryPoint Name="entryPoint1" Id="EntryPoint-8D5B661C9B89CB5">
      <AliasName Value="Execute" />
      <ComponentName Value="&lt;No Instance&gt;" />
      <DisplayName Value="" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Automation.EntryPoint" />
      <InstanceUniqueId Value="Automator-8D5B661C6B58228\EntryPoint-8D5B661C9B89CB5" />
      <MemberDetails Value="" />
      <MethodName Value="_EntryPointExecute" />
      <Removing Value="False" />
      <UseAlias Value="True" />
      <Content Name="Controls">
        <Capacity Value="4" />
        <Items>
          <OpenSpan.Automation.HiddenTypeProxy Name="hiddenTypeProxy1" Id="HiddenTypeProxy-8D5B661D7F2074B">
            <AliasName Value="param1" />
            <Parent Value="ComponentReference" Name="entryPoint1" />
            <ProxiedTypeName Value="System.Data.DataRow, System.Data" />
            <Scope Value="Local" Extended="True" />
            <UseAlias Value="True" />
            <Content Name="DynamicMembers">
              <Items>
                <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Instance" canRead="True" canWrite="True" type="System.Data.DataRow" aliasName="Instance" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
              </Items>
            </Content>
          </OpenSpan.Automation.HiddenTypeProxy>
        </Items>
      </Content>
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Result" canRead="True" canWrite="True" type="System.Data.DataRow" aliasName="Result" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="param1" canRead="False" canWrite="True" type="System.Data.DataRow" aliasName="param1" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Automation.EntryPoint>
    <OpenSpan.Automation.ExitPoint Name="exitPoint1" Id="ExitPoint-8D5B66247290BA4">
      <ComponentName Value="Execute" />
      <DisplayName Value="Exit1" />
      <EntryPoint Value="ComponentReference" Name="entryPoint1" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Automation.EntryPoint" />
      <InstanceUniqueId Value="Automator-8D5B661C6B58228\EntryPoint-8D5B661C9B89CB5" />
      <MemberDetails Value="" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Result" canRead="True" canWrite="True" type="System.Data.DataRow" aliasName="Result" shouldSerialize="False" visibility="DefaultOn" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Automation.ExitPoint>
  </Component>
</OpenSpanDesignDocument>