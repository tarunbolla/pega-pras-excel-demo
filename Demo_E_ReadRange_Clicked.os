<OpenSpanDesignDocument Version="8.0.1026.0" Serializer="2.0" Culture="en-US">
  <ComponentInfo>
    <Type Value="OpenSpan.Automation.Automator" />
    <Assembly Value="OpenSpan.Automation" />
    <AssemblyReferences>
      <Assembly Value="System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" />
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
    <OpenSpan.Automation.Automator Name="Demo_E_ReadRange_Clicked" Id="Automator-8D53B348BCA7280">
      <AutomationDocument>
        <Name Value="" />
        <Size Value="5000, 5000" />
        <Objects>
          <ConnectionBlock>
            <DisplayName Value="Control.Click" />
            <ConnectableUniqueId Value="Automator-8D53B348BCA7280\ConnectableEvent-8D53B349555AF40" />
            <PartID Value="1" />
            <Left Value="40" />
            <Top Value="60" />
            <Collapsed Value="True" />
            <WillExecute Value="True" />
            <InstanceName Value="btnReadRange" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock type="OpenSpan.Automation.Design.ConnectionBlocks.EntryPointExecuteBlock">
            <DisplayName Value="Execute" />
            <ConnectableUniqueId Value="Automator-8D53B348BCA7280\ConnectableMethod-8D53B3544D92C30" />
            <PartID Value="6" />
            <Left Value="240" />
            <Top Value="60" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="Demo_P_ReturnDataTableForRange" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D53B348BCA7280\ConnectableProperties-8D53B354B21E820" />
            <PartID Value="8" />
            <Left Value="20" />
            <Top Value="140" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="txtRangeFrom" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D53B348BCA7280\ConnectableProperties-8D53B354E4B89C0" />
            <PartID Value="9" />
            <Left Value="60" />
            <Top Value="220" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="txtRangeTo" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="Properties" />
            <ConnectableUniqueId Value="Automator-8D53B348BCA7280\ConnectableProperties-8D53B35565A5FB0" />
            <PartID Value="12" />
            <Left Value="620" />
            <Top Value="60" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="dataSheet" />
            <OverriddenIds />
          </ConnectionBlock>
        </Objects>
        <Links>
          <Link PartID="7" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="1" PortName="Raised" PortType="Event" ConnectableId="Automator-8D53B348BCA7280\ConnectableEvent-8D53B349555AF40" MemberComponentId="Automator-8D53B348BCA7280\EMPTY" />
            <To PartID="6" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53B348BCA7280\ConnectableMethod-8D53B3544D92C30" MemberComponentId="Automator-8D53B348BCA7280\ConnectableMethod-8D53B3544D92C30" />
            <LinkPoints>
              <Point value="181, 89" />
              <Point value="191, 89" />
              <Point value="213, 89" />
              <Point value="213, 89" />
              <Point value="235, 89" />
              <Point value="245, 89" />
            </LinkPoints>
          </Link>
          <Link PartID="10" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="8" PortName="Text" PortType="Property" ConnectableId="Automator-8D53B348BCA7280\ConnectableProperties-8D53B354B21E820" MemberComponentId="DesignForm-8D53AA1BDDF1678\TextBox-8D53B33FC4125A0" />
            <To PartID="6" PortName="param1" PortType="Property" ConnectableId="Automator-8D53B348BCA7280\ConnectableMethod-8D53B3544D92C30" MemberComponentId="Automator-8D53B348BCA7280\ConnectableMethod-8D53B3544D92C30" />
            <LinkPoints>
              <Point value="156, 186" />
              <Point value="166, 186" />
              <Point value="172, 186" />
              <Point value="172, 123" />
              <Point value="235, 123" />
              <Point value="245, 123" />
            </LinkPoints>
          </Link>
          <Link PartID="11" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="9" PortName="Text" PortType="Property" ConnectableId="Automator-8D53B348BCA7280\ConnectableProperties-8D53B354E4B89C0" MemberComponentId="DesignForm-8D53AA1BDDF1678\TextBox-8D53B33FF4E65A0" />
            <To PartID="6" PortName="param2" PortType="Property" ConnectableId="Automator-8D53B348BCA7280\ConnectableMethod-8D53B3544D92C30" MemberComponentId="Automator-8D53B348BCA7280\ConnectableMethod-8D53B3544D92C30" />
            <LinkPoints>
              <Point value="180, 266" />
              <Point value="190, 266" />
              <Point value="196, 266" />
              <Point value="196, 140" />
              <Point value="235, 140" />
              <Point value="245, 140" />
            </LinkPoints>
          </Link>
          <Link PartID="13" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="6" PortName="Complete" PortType="Event" ConnectableId="Automator-8D53B348BCA7280\ConnectableMethod-8D53B3544D92C30" MemberComponentId="Automator-8D53B34D7853250\ExitPoint-8D53B34EDF49AD0" />
            <To PartID="12" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53B348BCA7280\ConnectableProperties-8D53B35565A5FB0" MemberComponentId="Automator-8D53B348BCA7280\ConnectableProperties-8D53B35565A5FB0" />
            <LinkPoints>
              <Point value="516, 106" />
              <Point value="526, 106" />
              <Point value="570, 106" />
              <Point value="570, 89" />
              <Point value="615, 89" />
              <Point value="625, 89" />
            </LinkPoints>
          </Link>
          <Link PartID="14" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="6" PortName="Result" PortType="Property" ConnectableId="Automator-8D53B348BCA7280\ConnectableMethod-8D53B3544D92C30" MemberComponentId="Automator-8D53B348BCA7280\ConnectableMethod-8D53B3544D92C30" />
            <To PartID="12" PortName="DataSource" PortType="Property" ConnectableId="Automator-8D53B348BCA7280\ConnectableProperties-8D53B35565A5FB0" MemberComponentId="DesignForm-8D53AA1BDDF1678\DataGridView-8D53AA1F6ACCB63" />
            <LinkPoints>
              <Point value="516, 157" />
              <Point value="526, 157" />
              <Point value="570, 157" />
              <Point value="570, 106" />
              <Point value="615, 106" />
              <Point value="625, 106" />
            </LinkPoints>
          </Link>
        </Links>
        <Comments />
        <SubGraphs />
      </AutomationDocument>
    </OpenSpan.Automation.Automator>
    <OpenSpan.Automation.ConnectableEvent Name="connectableEvent1" Id="ConnectableEvent-8D53B349555AF40">
      <ComponentName Value="btnReadRange" />
      <DisplayName Value="Control.Click" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="System.Windows.Forms.Button" />
      <InstanceUniqueId Value="DesignForm-8D53AA1BDDF1678\Button-8D53B340760E510" />
      <MemberDetails Value=".Click Event" />
      <Content Name="MemberPrototypes">
        <Items>
          <OpenSpan.Automation.MemberPrototype>
            <MemberName Value="Click" />
            <MemberType Value="Event" />
            <TypeName Value="System.EventHandler" />
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.ConnectableEvent>
    <OpenSpan.Automation.ConnectableMethod Name="connectableMethod1" Id="ConnectableMethod-8D53B3544D92C30">
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
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties1" Id="ConnectableProperties-8D53B354B21E820">
      <ComponentName Value="txtRangeFrom" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="System.Windows.Forms.TextBox" />
      <InstanceUniqueId Value="DesignForm-8D53AA1BDDF1678\TextBox-8D53B33FC4125A0" />
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
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties2" Id="ConnectableProperties-8D53B354E4B89C0">
      <ComponentName Value="txtRangeTo" />
      <DefaultValues Value="" />
      <DisplayName Value="Properties" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="System.Windows.Forms.TextBox" />
      <InstanceUniqueId Value="DesignForm-8D53AA1BDDF1678\TextBox-8D53B33FF4E65A0" />
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
    <OpenSpan.Automation.ConnectableProperties Name="connectableProperties3" Id="ConnectableProperties-8D53B35565A5FB0">
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
  </Component>
</OpenSpanDesignDocument>