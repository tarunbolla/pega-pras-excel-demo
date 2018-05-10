<OpenSpanDesignDocument Version="8.0.1026.0" Serializer="2.0" Culture="en-US">
  <ComponentInfo>
    <Type Value="OpenSpan.Automation.Automator" />
    <Assembly Value="OpenSpan.Automation" />
    <AssemblyReferences>
      <Assembly Value="mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" />
      <Assembly Value="OpenSpan, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" />
      <Assembly Value="OpenSpan.Automation, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" />
      <Assembly Value="OpenSpan.Office, Version=8.0.0.0, Culture=neutral, PublicKeyToken=f5db91edc02d8c5e" />
      <Assembly Value="System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" />
    </AssemblyReferences>
    <DynamicAssemblyReferences />
    <FileReferences />
    <BuildReferences />
  </ComponentInfo>
  <Component Version="1.0">
    <OpenSpan.Automation.Automator Name="Demo_P_ReturnDataTableForRange" Id="Automator-8D53B34D7853250">
      <AutomationDocument>
        <Name Value="" />
        <Size Value="5011, 5000" />
        <Objects>
          <ConnectionBlock>
            <ConnectableUniqueId Value="Automator-8D53B34D7853250\EntryPoint-8D53B34E1E99100" />
            <Left Value="60" />
            <Top Value="140" />
            <PartID Value="1" />
          </ConnectionBlock>
          <ConnectionBlock type="OpenSpan.Automation.Design.ConnectionBlocks.MultiExitPointBlock">
            <ConnectableUniqueId Value="Automator-8D53B34D7853250\ExitPoint-8D53B34EDF49AD0" />
            <Left Value="420" />
            <Top Value="20" />
            <PartID Value="2" />
            <Title Value="Exit" />
            <EventName Value="" />
          </ConnectionBlock>
          <ConnectionBlock>
            <ConnectableUniqueId Value="Automator-8D53B34D7853250\LabelHost-8D53B350DA7AFC0" />
            <Left Value="240" />
            <Top Value="20" />
            <PartID Value="3" />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="ExportData" />
            <ConnectableUniqueId Value="Automator-8D53B34D7853250\ConnectableMethod-8D53B3527022A40" />
            <PartID Value="6" />
            <Left Value="260" />
            <Top Value="140" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="microsoftExcel1" />
            <OverriddenIds />
          </ConnectionBlock>
          <ConnectionBlock>
            <DisplayName Value="GoToLabel" />
            <ConnectableUniqueId Value="Automator-8D53B34D7853250\JumpHost-8D53B3532E33DE0" />
            <PartID Value="10" />
            <Left Value="480" />
            <Top Value="140" />
            <Collapsed Value="False" />
            <WillExecute Value="True" />
            <InstanceName Value="Jump To" />
            <OverriddenIds />
          </ConnectionBlock>
        </Objects>
        <Links>
          <Link PartID="4" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="3" PortName="_param1" PortType="Property" ConnectableId="Automator-8D53B34D7853250\LabelHost-8D53B350DA7AFC0" MemberComponentId="EMPTY" />
            <To PartID="2" PortName="Result" PortType="Property" ConnectableId="Automator-8D53B34D7853250\ExitPoint-8D53B34EDF49AD0" MemberComponentId="EMPTY" />
            <LinkPoints>
              <Point value="361, 65" />
              <Point value="371, 65" />
              <Point value="391, 65" />
              <Point value="391, 65" />
              <Point value="412, 65" />
              <Point value="422, 65" />
            </LinkPoints>
          </Link>
          <Link PartID="5" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="3" PortName="Complete" PortType="Event" ConnectableId="Automator-8D53B34D7853250\LabelHost-8D53B350DA7AFC0" MemberComponentId="Automator-8D53B34D7853250\LabelHost-8D53B350DA7AFC0" />
            <To PartID="2" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53B34D7853250\ExitPoint-8D53B34EDF49AD0" MemberComponentId="Automator-8D53B34D7853250\ExitPoint-8D53B34EDF49AD0" />
            <LinkPoints>
              <Point value="361, 38" />
              <Point value="371, 38" />
              <Point value="391, 38" />
              <Point value="391, 38" />
              <Point value="412, 38" />
              <Point value="422, 38" />
            </LinkPoints>
          </Link>
          <Link PartID="7" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="1" PortName="Complete" PortType="Event" ConnectableId="Automator-8D53B34D7853250\EntryPoint-8D53B34E1E99100" MemberComponentId="Automator-8D53B34D7853250\EntryPoint-8D53B34E1E99100" />
            <To PartID="6" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53B34D7853250\ConnectableMethod-8D53B3527022A40" MemberComponentId="Automator-8D53B34D7853250\ConnectableMethod-8D53B3527022A40" />
            <LinkPoints>
              <Point value="175, 158" />
              <Point value="185, 158" />
              <Point value="220, 158" />
              <Point value="220, 169" />
              <Point value="255, 169" />
              <Point value="265, 169" />
            </LinkPoints>
          </Link>
          <Link PartID="8" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="1" PortName="param1" PortType="Property" ConnectableId="Automator-8D53B34D7853250\EntryPoint-8D53B34E1E99100" MemberComponentId="EMPTY" />
            <To PartID="6" PortName="cellStart" PortType="Property" ConnectableId="Automator-8D53B34D7853250\ConnectableMethod-8D53B3527022A40" MemberComponentId="Automator-8D53B34D7853250\ConnectableMethod-8D53B3527022A40" />
            <LinkPoints>
              <Point value="175, 185" />
              <Point value="185, 185" />
              <Point value="185, 186" />
              <Point value="185, 186" />
              <Point value="255, 186" />
              <Point value="265, 186" />
            </LinkPoints>
          </Link>
          <Link PartID="9" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="1" PortName="param2" PortType="Property" ConnectableId="Automator-8D53B34D7853250\EntryPoint-8D53B34E1E99100" MemberComponentId="EMPTY" />
            <To PartID="6" PortName="cellEnd" PortType="Property" ConnectableId="Automator-8D53B34D7853250\ConnectableMethod-8D53B3527022A40" MemberComponentId="Automator-8D53B34D7853250\ConnectableMethod-8D53B3527022A40" />
            <LinkPoints>
              <Point value="175, 201" />
              <Point value="185, 201" />
              <Point value="185, 203" />
              <Point value="185, 203" />
              <Point value="255, 203" />
              <Point value="265, 203" />
            </LinkPoints>
          </Link>
          <Link PartID="11" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="6" PortName="Complete" PortType="Event" ConnectableId="Automator-8D53B34D7853250\ConnectableMethod-8D53B3527022A40" MemberComponentId="Automator-8D53B34D7853250\ConnectableMethod-8D53B3527022A40" />
            <To PartID="10" PortName="DoWork" PortType="Method" ConnectableId="Automator-8D53B34D7853250\JumpHost-8D53B3532E33DE0" MemberComponentId="Automator-8D53B34D7853250\JumpHost-8D53B3532E33DE0" />
            <LinkPoints>
              <Point value="406, 169" />
              <Point value="416, 169" />
              <Point value="420, 169" />
              <Point value="420, 157" />
              <Point value="473, 157" />
              <Point value="483, 157" />
            </LinkPoints>
          </Link>
          <Link PartID="12" Sensitive="False" Asynchronous="False" LogBeforeExecution="" LogAfterExecution="">
            <From PartID="6" PortName="Result" PortType="Property" ConnectableId="Automator-8D53B34D7853250\ConnectableMethod-8D53B3527022A40" MemberComponentId="Automator-8D53B34D7853250\ConnectableMethod-8D53B3527022A40" />
            <To PartID="10" PortName="_param1" PortType="Property" ConnectableId="Automator-8D53B34D7853250\JumpHost-8D53B3532E33DE0" MemberComponentId="Automator-8D53B34D7853250\JumpHost-8D53B3532E33DE0" />
            <LinkPoints>
              <Point value="406, 237" />
              <Point value="416, 237" />
              <Point value="446, 237" />
              <Point value="446, 186" />
              <Point value="475, 186" />
              <Point value="485, 186" />
            </LinkPoints>
          </Link>
        </Links>
        <Comments />
        <SubGraphs />
      </AutomationDocument>
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicMethodInfo dynamicType="Method" name="_EntryPointExecute" aliasName="Execute" visibility="DefaultOn" source="" blockTypeName="OpenSpan.Automation.Design.ConnectionBlocks.EntryPointExecuteBlock" returnType="System.Data.DataTable">
            <param name="param1" aliasName="fromRange" paramType="System.String" isIn="True" isOut="False" position="0" />
            <param name="param2" aliasName="toRange" paramType="System.String" isIn="True" isOut="False" position="1" />
          </OpenSpan.DynamicMembers.DynamicMethodInfo>
        </Items>
      </Content>
    </OpenSpan.Automation.Automator>
    <OpenSpan.Automation.EntryPoint Name="entryPoint1" Id="EntryPoint-8D53B34E1E99100">
      <AliasName Value="Execute" />
      <ComponentName Value="&lt;No Instance&gt;" />
      <DisplayName Value="" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Automation.EntryPoint" />
      <InstanceUniqueId Value="Automator-8D53B34D7853250\EntryPoint-8D53B34E1E99100" />
      <MemberDetails Value="" />
      <MethodName Value="_EntryPointExecute" />
      <Removing Value="False" />
      <UseAlias Value="True" />
      <Content Name="Controls">
        <Capacity Value="4" />
        <Items>
          <OpenSpan.Automation.HiddenTypeProxy Name="hiddenTypeProxy1" Id="HiddenTypeProxy-8D53B34E30D9EA0">
            <AliasName Value="fromRange" />
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
          <OpenSpan.Automation.HiddenTypeProxy Name="hiddenTypeProxy2" Id="HiddenTypeProxy-8D53B34E36AD9D0">
            <AliasName Value="toRange" />
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
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Result" canRead="True" canWrite="True" type="System.Data.DataTable" aliasName="Result" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="param1" canRead="False" canWrite="True" type="System.String" aliasName="fromRange" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="param2" canRead="False" canWrite="True" type="System.String" aliasName="toRange" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Automation.EntryPoint>
    <OpenSpan.Automation.ExitPoint Name="exitPoint1" Id="ExitPoint-8D53B34EDF49AD0">
      <ComponentName Value="Execute" />
      <DisplayName Value="Exit" />
      <EntryPoint Value="ComponentReference" Name="entryPoint1" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Automation.EntryPoint" />
      <InstanceUniqueId Value="Automator-8D53B34D7853250\EntryPoint-8D53B34E1E99100" />
      <MemberDetails Value="" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="Result" canRead="True" canWrite="True" type="System.Data.DataTable" aliasName="Result" shouldSerialize="False" visibility="DefaultOn" source="" blockTypeName="" />
        </Items>
      </Content>
    </OpenSpan.Automation.ExitPoint>
    <OpenSpan.Automation.LabelHost Name="labelHost1" Id="LabelHost-8D53B350DA7AFC0">
      <ComponentName Value="OpenSpan.Automation.EntryPoint" />
      <DisplayName Value="Exit" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Automation.EntryPoint" />
      <InstanceUniqueId Value="EMPTY" />
      <LabelName Value="Exit" />
      <MemberDetails Value="" />
      <Content Name="DynamicMembers">
        <Items>
          <OpenSpan.DynamicMembers.DynamicPropertyInfo dynamicType="Property" name="_param1" canRead="True" canWrite="True" type="System.Data.DataTable" aliasName="param1" shouldSerialize="False" visibility="AlwaysHidden" source="" blockTypeName="" />
          <OpenSpan.DynamicMembers.DynamicMethodInfo dynamicType="Method" name="GoToLabel" aliasName="GoToLabel" visibility="AlwaysHidden" source="" blockTypeName="" returnType="System.Void">
            <param name="_param1" aliasName="param1" paramType="System.Data.DataTable" isIn="True" isOut="False" position="0" />
          </OpenSpan.DynamicMembers.DynamicMethodInfo>
        </Items>
      </Content>
    </OpenSpan.Automation.LabelHost>
    <OpenSpan.Automation.ConnectableMethod Name="connectableMethod1" Id="ConnectableMethod-8D53B3527022A40">
      <ComponentName Value="microsoftExcel1" />
      <DisplayName Value="ExportData" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Office.MicrosoftExcel" />
      <InstanceUniqueId Value="GlobalContainer-8D53AA1A011B17D\MicrosoftExcel-8D53AA1A49B9D1B" />
      <MemberDetails Value=".ExportData() Method" />
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
            <MemberName Value="ExportData" />
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
                      <ParamName Value="cellStart" />
                      <Position Value="0" />
                      <TypeName Value="System.String" />
                    </OpenSpan.Automation.ParameterPrototype>
                    <OpenSpan.Automation.ParameterPrototype>
                      <CanRead Value="False" />
                      <CanWrite Value="True" />
                      <DefaultSet Value="False" />
                      <ParamName Value="cellEnd" />
                      <Position Value="1" />
                      <TypeName Value="System.String" />
                    </OpenSpan.Automation.ParameterPrototype>
                    <OpenSpan.Automation.ParameterPrototype>
                      <CanRead Value="False" />
                      <CanWrite Value="True" />
                      <DefaultSet Value="True" />
                      <DefaultValue Value="True" />
                      <ParamName Value="createHeader" />
                      <Position Value="2" />
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
    <OpenSpan.Automation.JumpHost Name="jumpHost1" Id="JumpHost-8D53B3532E33DE0">
      <ComponentName Value="labelHost1" />
      <DisplayName Value="GoToLabel" />
      <ExceptionsHandled Value="False" />
      <InstanceTypeName Value="OpenSpan.Automation.LabelHost" />
      <InstanceUniqueId Value="Automator-8D53B34D7853250\LabelHost-8D53B350DA7AFC0" />
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
                      <TypeAssemblyName Value="System.Data" />
                      <TypeName Value="System.Data.DataTable" />
                    </OpenSpan.Automation.ParameterPrototype>
                  </Items>
                </Content>
              </OpenSpan.Automation.MethodSignature>
            </Content>
          </OpenSpan.Automation.MemberPrototype>
        </Items>
      </Content>
    </OpenSpan.Automation.JumpHost>
  </Component>
</OpenSpanDesignDocument>