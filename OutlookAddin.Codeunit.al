codeunit 50100 "Office Handler BV"
{
    TableNo = "Office Add-in Context";

    trigger OnRun()
    begin
        HandleCommand(Rec);
    end;

    var
        AppIdTxt: Label 'a5fcb2f3-fce3-47a5-976a-e837599ae46d', Locked = true;
        AddinNameTxt: Label '4PS Construct Insights';
        AddinDescriptionTxt: Label 'Provides 4PS Contruct information directly within Outlook messages.';
        BrandingFolderTxt: Label 'ProjectMadeira', Locked = true;
        AddinManifestManagement: Codeunit "Add-in Manifest Management";
        EnvironmentInfo: Codeunit "Environment Information";

    internal procedure GetManifestVersion(): Text
    begin
        exit('2.0.0.0');
    end;

    local procedure HandleCommand(TempOfficeAddinContext: Record "Office Add-in Context" temporary)
    var
        CustomOutlookAction: Record "Custom Outlook Action";
    begin
        case TempOfficeAddinContext.Command of
            'taskPaneUrl':
                Codeunit.Run(Codeunit::"Office Contact Handler", TempOfficeAddinContext);
        end;

        if CustomOutlookAction.FindSet then
            repeat
                if TempOfficeAddinContext.Command = CustomOutlookAction.GetCommand then
                    OpenPage(CustomOutlookAction, TempOfficeAddinContext);
            until CustomOutlookAction.Next = 0;
    end;

    local procedure OpenPage(CustomOutlookAction: Record "Custom Outlook Action"; TempOfficeAddinContext: Record "Office Add-in Context")
    var
        RecRef: RecordRef;
        VarRecRef: Variant;
        Customer: Record Customer;
        Vendor: Record Vendor;
        Contact: Record Contact;
        Insert: Boolean;
    begin
        Insert := false;
        if CustomOutlookAction."Table ID" <> 0 then begin
            RecRef.Open(CustomOutlookAction."Table ID");
            RecRef.Init();


            Case CustomOutlookAction."Contact Type" of
                CustomOutlookAction."Contact Type"::Vendor:
                    begin
                        Vendor.SetRange("E-Mail", TempOfficeAddinContext.Email);
                        if Vendor.FindFirst then begin
                            RecRef.Field(CustomOutlookAction."Field No. Contact Type").Value := Vendor."No.";
                            Insert := true;
                        end;
                    end;
                CustomOutlookAction."Contact Type"::Customer:
                    begin
                        Customer.SetRange("E-Mail", TempOfficeAddinContext.Email);
                        if Customer.FindFirst then begin
                            RecRef.Field(CustomOutlookAction."Field No. Contact Type").Value := Customer."No.";
                            Insert := true;
                        end;
                    end;
                CustomOutlookAction."Contact Type"::Contact:
                    begin
                        Contact.SetRange("E-Mail", TempOfficeAddinContext.Email);
                        if Contact.FindFirst then begin
                            RecRef.Field(CustomOutlookAction."Field No. Contact Type").Value := Contact."No.";
                            Insert := true;
                        end;
                    end;
            end;

            if CustomOutlookAction."Field No. Name" <> 0 then begin
                RecRef.Field(CustomOutlookAction."Field No. Name").Value := TempOfficeAddinContext.Name;
                Insert := true;
            end;
            if CustomOutlookAction."Field No. E-mail" <> 0 then begin
                RecRef.Field(CustomOutlookAction."Field No. E-mail").Value := TempOfficeAddinContext.Email;
                Insert := true;
            end;

            if Insert then begin
                RecRef.Insert(true);
                Commit;
                VarRecRef := RecRef;
                PAGE.Run(CustomOutlookAction."Page ID", VarRecRef);
            end;

        end else
            PAGE.Run(CustomOutlookAction."Page ID");
    end;

    local procedure CanHandle(CodeunitID: Integer): Boolean
    begin
        exit(CodeunitID = CODEUNIT::"Office Handler BV");
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Office Management", 'OnGetExternalHandlerCodeunit', '', true, false)]
    local procedure RegisterExternalHandler(OfficeAddinContext: Record "Office Add-in Context"; HostType: Text; var HandlerCodeunit: Integer)
    var
        CustomOutlookNewAction: Record "Custom Outlook Action";
    begin
        if HostType = 'OutlookTaskPane' then
            HandlerCodeunit := CODEUNIT::"Office Handler BV";

        if CustomOutlookNewAction.FindSet then
            repeat
                if OfficeAddinContext.Command = CustomOutlookNewAction.GetCommand then
                    HandlerCodeunit := CODEUNIT::"Office Handler BV";
            until CustomOutlookNewAction.Next = 0;
    end;

    local procedure WebClientUrl() BaseURL: Text
    begin
        BaseURL := GetUrl(CLIENTTYPE::Web);
        if EnvironmentInfo.IsSaaS() then
            BaseURL := StrSubstNo('%1/%2', BaseURL, BrandingFolderTxt);
        exit(BaseURL);
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Add-in Manifest Management", 'CreateDefaultAddins', '', false, false)]
    local procedure OnCreateAddin(var OfficeAddin: Record "Office Add-in")
    begin
        if OfficeAddin.Get('a5fcb2f3-fce3-47a5-976a-e837599ae46d') then
            OfficeAddin.Delete();

        OfficeAddin.Init();
        OfficeAddin."Application ID" := AppIdTxt;
        OfficeAddin."Manifest Codeunit" := CODEUNIT::"Office Handler BV";
        OfficeAddin.Name := AddinNameTxt;
        OfficeAddin.Description := AddinDescriptionTxt;
        OfficeAddin.Version := GetManifestVersion();
        OfficeAddin.Insert(true);

        OfficeAddin.SetDefaultManifestText(DefaultManifestText());
        OfficeAddin.Modify(true);
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Add-in Manifest Management", 'GetAddin', '', false, false)]
    local procedure OnGetAddin(var OfficeAddin: Record "Office Add-in"; CodeunitID: Integer)
    begin
        if CodeunitID = CODEUNIT::"Office Handler BV" then
            OfficeAddin.Get(AppIdTxt);
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Add-in Manifest Management", 'GetAddinID', '', false, false)]
    local procedure OnGetAddinID(var ID: Text; CodeunitID: Integer)
    begin
        if CodeunitID = CODEUNIT::"Office Handler BV" then
            ID := AppIdTxt;
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Add-in Manifest Management", 'GetAddinVersion', '', false, false)]
    local procedure OnGetAddinVersion(var Version: Text; CodeunitID: Integer)
    begin
        if CodeunitID = CODEUNIT::"Office Handler BV" then
            Version := GetManifestVersion();
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Add-in Manifest Management", 'GetManifestCodeunit', '', false, false)]
    local procedure OnGetCodeunitID(var CodeunitID: Integer; HostType: Text)
    begin
        if HostType in ['OutlookTaskPane'] then
            CodeunitID := CODEUNIT::"Office Handler BV";
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Add-in Manifest Management", 'OnGenerateManifest', '', false, false)]
    local procedure OnGenerateManifest(var OfficeAddin: Record "Office Add-in"; var ManifestText: Text; CodeunitID: Integer)
    begin
        if not CanHandle(CodeunitID) then
            exit;

        ManifestText := OfficeAddin.GetDefaultManifestText();
    end;

    local procedure DefaultManifestText() Value: Text
    var
        CustomOutlookNewAction: Record "Custom Outlook Action";
        i: Integer;
    begin
        Value :=
          '<?xml version="1.0" encoding="utf-8"?>' +
          '<OfficeApp' +
          '  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"' +
          '  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
          '  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"' +
          '  xmlns:o10="http://schemas.microsoft.com/office/mailappversionoverrides"' +
          '  xmlns:o11="http://schemas.microsoft.com/office/mailappversionoverrides/1.1"' +
          '  xsi:type="MailApp">' +
          '  <Id>' + AppIdTxt + '</Id>' +
          '  <Version>' + GetManifestVersion() + '</Version>' +
          '  <ProviderName>Microsoft</ProviderName>' +
          '  <DefaultLocale>en-US</DefaultLocale>' +
          '  <DisplayName DefaultValue="4PS Construct Insights" />' +
          '  <Description DefaultValue="' + AddinDescriptionTxt + '" />' +
          '  <IconUrl DefaultValue="' + AddinManifestManagement.XMLEncode(GetUrl(CLIENTTYPE::Web)) + '/Resources/Images/';
        if EnvironmentInfo.IsSaaS() then
            Value := Value + BrandingFolderTxt + '/OfficeAddin_64x.png"/>'
        else
            Value := Value + '/OfficeAddin_64x.png"/>';

        Value := Value +
        '  <HighResolutionIconUrl DefaultValue="' + AddinManifestManagement.XMLEncode(GetUrl(CLIENTTYPE::Web)) + '/Resources/Images/';
        if EnvironmentInfo.IsSaaS() then
            Value := Value + BrandingFolderTxt + '/OfficeAddin_64x.png"/>'
        else
            Value := Value + 'OfficeAddin_64x.png"/>';

        Value := Value +
        '  <AppDomains>' +
        '    <AppDomain>' + AddinManifestManagement.XMLEncode(GetUrl(CLIENTTYPE::Web)) + '</AppDomain>' +
        '  </AppDomains>' +
        '  <Hosts>' +
        '    <Host Name="Mailbox" />' +
        '  </Hosts>' +
        '  <Requirements>' +
        '    <Sets>' +
        '      <Set Name="MailBox" MinVersion="1.3" />' +
        '    </Sets>' +
        '  </Requirements>' +
        '  <FormSettings>' +
        '    <Form xsi:type="ItemRead">' +
        '      <DesktopSettings>' +
        '        <SourceLocation DefaultValue="' + AddinManifestManagement.XMLEncode(addinManifestManagement.ConstructURL('ItemRead', '', GetManifestVersion())) + '" />' +
        '        <RequestedHeight>300</RequestedHeight>' +
        '      </DesktopSettings>' +
        '      <TabletSettings>' +
        '        <SourceLocation DefaultValue="' + AddinManifestManagement.XMLEncode(addinManifestManagement.ConstructURL('ItemRead', '', GetManifestVersion())) + '" />' +
        '        <RequestedHeight>400</RequestedHeight>' +
        '      </TabletSettings>' +
        '      <PhoneSettings>' +
        '        <SourceLocation DefaultValue="' + AddinManifestManagement.XMLEncode(addinManifestManagement.ConstructURL('ItemRead', '', GetManifestVersion())) + AddinManifestManagement.XMLEncode('&isphone=1') + '" />' +
        '      </PhoneSettings>' +
        '    </Form>' +
        '    <Form xsi:type="ItemEdit">' +
        '      <DesktopSettings>' +
        '        <SourceLocation DefaultValue="' + AddinManifestManagement.XMLEncode(addinManifestManagement.ConstructURL('ItemEdit', '', GetManifestVersion())) + '" />' +
        '      </DesktopSettings>' +
        '      <TabletSettings>' +
        '        <SourceLocation DefaultValue="' + AddinManifestManagement.XMLEncode(addinManifestManagement.ConstructURL('ItemEdit', '', GetManifestVersion())) + '" />' +
        '      </TabletSettings>' +
        '      <PhoneSettings>' +
        '        <SourceLocation DefaultValue="' + AddinManifestManagement.XMLEncode(addinManifestManagement.ConstructURL('ItemEdit', '', GetManifestVersion())) + AddinManifestManagement.XMLEncode('&isphone=1') + '" />' +
        '      </PhoneSettings>' +
        '    </Form>' +
        '  </FormSettings>' +
        '  <Permissions>ReadWriteMailbox</Permissions>' +
        '  <Rule xsi:type="RuleCollection" Mode="Or">' +
        '    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />' +
        '    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />' +
        '    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />' +
        '    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />' +
        '  </Rule>' +
        '' +
        '  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides"' +
        ' xsi:type="VersionOverridesV1_0">' +
        '    <Requirements>' +
        '      <bt:Sets DefaultMinVersion="1.3">' +
        '        <bt:Set Name="Mailbox" />' +
        '      </bt:Sets>' +
        '    </Requirements>' +
        '    <Hosts>' +
        '      <Host xsi:type="MailHost">' +
        '        <DesktopFormFactor>' +
        '          <!-- Custom pane, only applies to read form -->' +
        '          <ExtensionPoint xsi:type="CustomPane">' +
        '            <RequestedHeight>300</RequestedHeight>' +
        '            <SourceLocation resid="taskPaneUrl"/>' +
        '            <!-- Change this Mode to Or to enable the custom pane -->' +
        '            <Rule xsi:type="RuleCollection" Mode="And">' +
        '              <Rule xsi:type="ItemIs" ItemType="Message"/>' +
        '              <Rule xsi:type="ItemIs" ItemType="AppointmentAttendee"/>' +
        '            </Rule>' +
        '          </ExtensionPoint>' +
        '' +
        '          <!-- Message read form -->' +
        '          <ExtensionPoint xsi:type="MessageReadCommandSurface">' +
        '            <OfficeTab id="TabDefault">' +
        '              <Group id="msgReadGroup">' +
        '                <Label resid="groupLabel" />' +
        '                <Tooltip resid="groupTooltip" />' +
        '' +
        '                <!-- Task pane button -->' +
        '                <Control xsi:type="Button" id="msgReadOpenPaneButton">' +
        '                  <Label resid="openPaneButtonLabel" />' +
        '                  <Tooltip resid="openPaneButtonTooltip" />' +
        '                  <Supertip>' +
        '                    <Title resid="openPaneSuperTipTitle" />' +
        '                    <Description resid="openPaneSuperTipDesc" />' +
        '                  </Supertip>' +
        '                  <Icon>' +
        '                    <bt:Image size="16" resid="nav-icon-16" />' +
        '                    <bt:Image size="32" resid="nav-icon-32" />' +
        '                    <bt:Image size="80" resid="nav-icon-80" />' +
        '                  </Icon>' +
        '                  <Action xsi:type="ShowTaskpane">' +
        '                    <SourceLocation resid="taskPaneUrl" />' +
        '                  </Action>' +
        '                </Control>' +
        '';
        CustomOutlookNewAction.Reset;
        CustomOutlookNewAction.SetFilter("Table ID", '<>%1', 0);
        if CustomOutlookNewAction.FindSet then begin
            Value := Value +
        '                <!-- New Page group -->' +
        '                <Control xsi:type="Menu" id="newMenuReadButton">' +
        '                  <Label resid="newMenuButtonLabel" />' +
        '                  <Tooltip resid="newMenuButtonTooltip" />' +
        '                  <Supertip>' +
        '                    <Title resid="newMenuSuperTipTitle" />' +
        '                    <Description resid="newMenuSuperTipDesc" />' +
        '                  </Supertip>' +
        '                  <Icon>' +
        '                    <bt:Image size="16" resid="new-document-16" />' +
        '                    <bt:Image size="32" resid="new-document-32" />' +
        '                    <bt:Image size="80" resid="new-document-80" />' +
        '                  </Icon>' +
        '                  <Items>';
            i := 1;
            CustomOutlookNewAction.Reset;
            CustomOutlookNewAction.SetFilter("Table ID", '<>%1', 0);
            if CustomOutlookNewAction.FindSet then
                repeat
                    Value := Value + CustomOutlookNewAction.GetMenuReadItem(i, 'ReadComposeNew');
                    i := i + 1;
                until CustomOutlookNewAction.Next = 0;

            Value := Value +
            '                  </Items>' +
            '                </Control>';
        end;
        CustomOutlookNewAction.Reset;
        CustomOutlookNewAction.SetRange("Table ID", 0);
        if CustomOutlookNewAction.FindSet then begin
            Value := Value +
            '                <!-- Open Page group -->' +
            '                <Control xsi:type="Menu" id="pageMenuReadButton">' +
            '                  <Label resid="pageMenuButtonLabel" />' +
            '                  <Tooltip resid="pageMenuButtonTooltip" />' +
            '                  <Supertip>' +
            '                    <Title resid="pageMenuSuperTipTitle" />' +
            '                    <Description resid="pageMenuSuperTipDesc" />' +
            '                  </Supertip>' +
            '                  <Icon>' +
            '                    <bt:Image size="16" resid="nav-icon-16" />' +
            '                    <bt:Image size="32" resid="nav-icon-32" />' +
            '                    <bt:Image size="80" resid="nav-icon-80" />' +
            '                  </Icon>' +
            '                  <Items>';
            i := 1;
            CustomOutlookNewAction.Reset;
            CustomOutlookNewAction.SetRange("Table ID", 0);
            if CustomOutlookNewAction.FindSet then
                repeat
                    Value := Value + CustomOutlookNewAction.GetMenuReadItem(i, 'ReadComposeOpen');
                    i := i + 1;
                until CustomOutlookNewAction.Next = 0;


            Value := Value + '  </Items>' +
            '                </Control>';
        end;
        Value := Value +
          '              </Group>' +
          '            </OfficeTab>' +
          '          </ExtensionPoint>' +
          '' +
          '          <!-- Message compose form -->' +
          '          <ExtensionPoint xsi:type="MessageComposeCommandSurface">' +
          '            <OfficeTab id="TabDefault">' +
          '              <Group id="msgComposeGroup">' +
          '                <Label resid="groupLabel" />' +
          '                <Tooltip resid="groupTooltip" />' +
          '' +
          '                <!-- Task pane button -->' +
          '                <Control xsi:type="Button" id="msgComposeOpenPaneButton">' +
          '                  <Label resid="openPaneButtonLabel" />' +
          '                  <Tooltip resid="openPaneButtonTooltip" />' +
          '                  <Supertip>' +
          '                    <Title resid="openPaneSuperTipTitle" />' +
          '                    <Description resid="openPaneSuperTipDesc" />' +
          '                  </Supertip>' +
          '                  <Icon>' +
          '                    <bt:Image size="16" resid="nav-icon-16" />' +
          '                    <bt:Image size="32" resid="nav-icon-32" />' +
          '                    <bt:Image size="80" resid="nav-icon-80" />' +
          '                  </Icon>' +
          '                  <Action xsi:type="ShowTaskpane">' +
          '                    <SourceLocation resid="taskPaneUrl" />' +
          '                  </Action>' +
          '                </Control>' +
          '';
        CustomOutlookNewAction.Reset;
        CustomOutlookNewAction.SetFilter("Table ID", '<>%1', 0);
        if CustomOutlookNewAction.FindSet then begin
            Value := Value +
        '                <!-- New Page group -->' +
        '                <Control xsi:type="Menu" id="newMenuComposeButton">' +
        '                  <Label resid="newMenuButtonLabel" />' +
        '                  <Tooltip resid="newMenuButtonTooltip" />' +
        '                  <Supertip>' +
        '                    <Title resid="newMenuSuperTipTitle" />' +
        '                    <Description resid="newMenuSuperTipDesc" />' +
        '                  </Supertip>' +
        '                  <Icon>' +
        '                    <bt:Image size="16" resid="new-document-16" />' +
        '                    <bt:Image size="32" resid="new-document-32" />' +
        '                    <bt:Image size="80" resid="new-document-80" />' +
        '                  </Icon>' +
        '                  <Items>';
            i := 1;
            CustomOutlookNewAction.Reset;
            CustomOutlookNewAction.SetFilter("Table ID", '<>%1', 0);
            if CustomOutlookNewAction.FindSet then
                repeat
                    Value := Value + CustomOutlookNewAction.GetMenuReadItem(i, 'ComposeNew');
                    i := i + 1;
                until CustomOutlookNewAction.Next = 0;

            Value := Value +
            '                  </Items>' +
            '                </Control>';
        end;
        CustomOutlookNewAction.Reset;
        CustomOutlookNewAction.SetRange("Table ID", 0);
        if CustomOutlookNewAction.FindSet then begin
            Value := Value +
          '                <!-- Open Page group -->' +
          '                <Control xsi:type="Menu" id="pageMenuComposeButton">' +
          '                  <Label resid="pageMenuButtonLabel" />' +
          '                  <Tooltip resid="pageMenuButtonTooltip" />' +
          '                  <Supertip>' +
          '                    <Title resid="pageMenuSuperTipTitle" />' +
          '                    <Description resid="pageMenuSuperTipDesc" />' +
          '                  </Supertip>' +
          '                  <Icon>' +
          '                    <bt:Image size="16" resid="nav-icon-16" />' +
          '                    <bt:Image size="32" resid="nav-icon-32" />' +
          '                    <bt:Image size="80" resid="nav-icon-80" />' +
          '                  </Icon>' +
          '                  <Items>';
            i := 1;
            CustomOutlookNewAction.Reset;
            CustomOutlookNewAction.SetRange("Table ID", 0);
            if CustomOutlookNewAction.FindSet then
                repeat
                    Value := Value + CustomOutlookNewAction.GetMenuReadItem(i, 'ComposeOpen');
                    i := i + 1;
                until CustomOutlookNewAction.Next = 0;
            Value := Value + '  </Items>' +
            '                </Control>';
        end;
        Value := Value +
          '              </Group>' +
          '            </OfficeTab>' +
          '          </ExtensionPoint>' +
          '' +
          '          <!-- Appointment organizer form -->' +
          '          <ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">' +
          '            <OfficeTab id="TabDefault">' +
          '              <Group id="apptOrganizerGroup">' +
          '                <Label resid="groupLabel" />' +
          '                <Tooltip resid="groupTooltip" />' +
          '                <!-- Task pane button -->' +
          '                <Control xsi:type="Button" id="apptOrganizerOpenPaneButton">' +
          '                  <Label resid="openPaneButtonLabel" />' +
          '                  <Tooltip resid="openPaneButtonTooltip" />' +
          '                  <Supertip>' +
          '                    <Title resid="openPaneSuperTipTitle" />' +
          '                    <Description resid="openPaneSuperTipDesc" />' +
          '                  </Supertip>' +
          '                  <Icon>' +
          '                    <bt:Image size="16" resid="nav-icon-16" />' +
          '                    <bt:Image size="32" resid="nav-icon-32" />' +
          '                    <bt:Image size="80" resid="nav-icon-80" />' +
          '                  </Icon>' +
          '                  <Action xsi:type="ShowTaskpane">' +
          '                    <SourceLocation resid="taskPaneUrl" />' +
          '                  </Action>' +
          '                </Control>';
        CustomOutlookNewAction.Reset;
        CustomOutlookNewAction.SetFilter("Table ID", '<>%1', 0);
        if CustomOutlookNewAction.FindSet then begin
            Value := Value +
        '                <!-- New Page group -->' +
        '                <Control xsi:type="Menu" id="newMenuapptOrganizerButton">' +
        '                  <Label resid="newMenuButtonLabel" />' +
        '                  <Tooltip resid="newMenuButtonTooltip" />' +
        '                  <Supertip>' +
        '                    <Title resid="newMenuSuperTipTitle" />' +
        '                    <Description resid="newMenuSuperTipDesc" />' +
        '                  </Supertip>' +
        '                  <Icon>' +
        '                    <bt:Image size="16" resid="new-document-16" />' +
        '                    <bt:Image size="32" resid="new-document-32" />' +
        '                    <bt:Image size="80" resid="new-document-80" />' +
        '                  </Icon>' +
        '                  <Items>';
            i := 1;
            CustomOutlookNewAction.Reset;
            CustomOutlookNewAction.SetFilter("Table ID", '<>%1', 0);
            if CustomOutlookNewAction.FindSet then
                repeat
                    Value := Value + CustomOutlookNewAction.GetMenuReadItem(i, 'apptOrganizerNew');
                    i := i + 1;
                until CustomOutlookNewAction.Next = 0;

            Value := Value +
            '                  </Items>' +
            '                </Control>';
        end;
        CustomOutlookNewAction.Reset;
        CustomOutlookNewAction.SetRange("Table ID", 0);
        if CustomOutlookNewAction.FindSet then begin
            Value := Value +
            '                <!-- 4PS Page group -->' +
            '                <Control xsi:type="Menu" id="pageMenuapptOrganizerButton2">' +
            '                  <Label resid="pageMenuButtonLabel" />' +
            '                  <Tooltip resid="pageMenuButtonTooltip" />' +
            '                  <Supertip>' +
            '                    <Title resid="pageMenuSuperTipTitle" />' +
            '                    <Description resid="pageMenuSuperTipDesc" />' +
            '                  </Supertip>' +
            '                  <Icon>' +
            '                    <bt:Image size="16" resid="nav-icon-16" />' +
            '                    <bt:Image size="32" resid="nav-icon-32" />' +
            '                    <bt:Image size="80" resid="nav-icon-80" />' +
            '                  </Icon>' +
            '                  <Items>';
            i := 1;
            CustomOutlookNewAction.Reset;
            CustomOutlookNewAction.SetRange("Table ID", 0);
            if CustomOutlookNewAction.FindSet then
                repeat
                    Value := Value + CustomOutlookNewAction.GetMenuReadItem(i, 'apptOrganizerOpen');
                    i := i + 1;
                until CustomOutlookNewAction.Next = 0;


            Value := Value + '  </Items>' +
            '                </Control>';
        end;
        Value := Value +
          '              </Group>' +
          '            </OfficeTab>' +
          '          </ExtensionPoint>' +
          '' +
          '          <!-- Appointment attendee form -->' +
          '          <ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">' +
          '            <OfficeTab id="TabDefault">' +
          '              <Group id="apptAttendeeGroup">' +
          '                <Label resid="groupLabel" />' +
          '                <Tooltip resid="groupTooltip" />' +
          '                <!-- Task pane button -->' +
          '                <Control xsi:type="Button" id="apptAttendeeOpenPaneButton">' +
          '                  <Label resid="openPaneButtonLabel" />' +
          '                  <Tooltip resid="openPaneButtonTooltip" />' +
          '                  <Supertip>' +
          '                    <Title resid="openPaneSuperTipTitle" />' +
          '                    <Description resid="openPaneSuperTipDesc" />' +
          '                  </Supertip>' +
          '                  <Icon>' +
          '                    <bt:Image size="16" resid="nav-icon-16" />' +
          '                    <bt:Image size="32" resid="nav-icon-32" />' +
          '                    <bt:Image size="80" resid="nav-icon-80" />' +
          '                  </Icon>' +
          '                  <Action xsi:type="ShowTaskpane">' +
          '                    <SourceLocation resid="taskPaneUrl" />' +
          '                  </Action>' +
          '                </Control>';
        CustomOutlookNewAction.Reset;
        CustomOutlookNewAction.SetFilter("Table ID", '<>%1', 0);
        if CustomOutlookNewAction.FindSet then begin
            Value := Value +
        '                <!-- New Page group -->' +
        '                <Control xsi:type="Menu" id="newMenuapptAttendeeButton">' +
        '                  <Label resid="newMenuButtonLabel" />' +
        '                  <Tooltip resid="newMenuButtonTooltip" />' +
        '                  <Supertip>' +
        '                    <Title resid="newMenuSuperTipTitle" />' +
        '                    <Description resid="newMenuSuperTipDesc" />' +
        '                  </Supertip>' +
        '                  <Icon>' +
        '                    <bt:Image size="16" resid="new-document-16" />' +
        '                    <bt:Image size="32" resid="new-document-32" />' +
        '                    <bt:Image size="80" resid="new-document-80" />' +
        '                  </Icon>' +
        '                  <Items>';
            i := 1;
            CustomOutlookNewAction.Reset;
            CustomOutlookNewAction.SetFilter("Table ID", '<>%1', 0);
            if CustomOutlookNewAction.FindSet then
                repeat
                    Value := Value + CustomOutlookNewAction.GetMenuReadItem(i, 'apptAttendeeNew');
                    i := i + 1;
                until CustomOutlookNewAction.Next = 0;

            Value := Value +
            '                  </Items>' +
            '                </Control>';
        end;
        CustomOutlookNewAction.Reset;
        CustomOutlookNewAction.SetRange("Table ID", 0);
        if CustomOutlookNewAction.FindSet then begin
            Value := Value +
            '                <!-- 4PS Page group -->' +
            '                <Control xsi:type="Menu" id="pageMenuapptAttendeeButton">' +
            '                  <Label resid="pageMenuButtonLabel" />' +
            '                  <Tooltip resid="pageMenuButtonTooltip" />' +
            '                  <Supertip>' +
            '                    <Title resid="pageMenuSuperTipTitle" />' +
            '                    <Description resid="pageMenuSuperTipDesc" />' +
            '                  </Supertip>' +
            '                  <Icon>' +
            '                    <bt:Image size="16" resid="nav-icon-16" />' +
            '                    <bt:Image size="32" resid="nav-icon-32" />' +
            '                    <bt:Image size="80" resid="nav-icon-80" />' +
            '                  </Icon>' +
            '                  <Items>';
            i := 1;
            CustomOutlookNewAction.Reset;
            CustomOutlookNewAction.SetRange("Table ID", 0);
            if CustomOutlookNewAction.FindSet then
                repeat
                    Value := Value + CustomOutlookNewAction.GetMenuReadItem(i, 'apptAttendeeOpen');
                    i := i + 1;
                until CustomOutlookNewAction.Next = 0;


            Value := Value + '  </Items>' +
            '                </Control>';
        end;
        Value := Value +
          '              </Group>' +
          '            </OfficeTab>' +
          '          </ExtensionPoint>' +
          '        </DesktopFormFactor>' +
          '      </Host>' +
          '    </Hosts>' +
          '    <Resources>' +
          '      <bt:Images>' +
          '        <!-- NAV icon -->' +
          '        <bt:Image id="nav-icon-16" DefaultValue="' + AddinManifestManagement.XMLEncode(WebClientUrl()) + '/Resources/Images/';
        if EnvironmentInfo.IsSaaS() then
            Value := Value + BrandingFolderTxt + '/OfficeAddin_16x.png"/>'
        else
            Value := Value + '/OfficeAddin_16x.png"/>';

        Value := Value +
          '        <bt:Image id="nav-icon-32" DefaultValue="' + AddinManifestManagement.XMLEncode(WebClientUrl()) + '/Resources/Images/';
        if EnvironmentInfo.IsSaaS() then
            Value := Value + BrandingFolderTxt + '/OfficeAddin_32x.png"/>'
        else
            Value := Value + '/OfficeAddin_32x.png"/>';

        Value := Value +
          '        <bt:Image id="nav-icon-80" DefaultValue="' + AddinManifestManagement.XMLEncode(WebClientUrl()) + '/Resources/Images/';
        if EnvironmentInfo.IsSaaS() then
            Value := Value + BrandingFolderTxt + '/OfficeAddin_80x.png"/>'
        else
            Value := Value + '/OfficeAddin_80x.png"/>';

        Value := Value +
          '' +
          '        <!-- New document icon -->' +
          '        <bt:Image id="new-document-16" DefaultValue="' + AddinManifestManagement.XMLEncode(WebClientUrl()) + '/Resources/Images/NewDocument_16x16.png"/>' +
          '        <bt:Image id="new-document-32" DefaultValue="' + AddinManifestManagement.XMLEncode(WebClientUrl()) + '/Resources/Images/NewDocument_32x32.png"/>' +
          '        <bt:Image id="new-document-80" DefaultValue="' + AddinManifestManagement.XMLEncode(WebClientUrl()) + '/Resources/Images/NewDocument_80x80.png"/>' +
          '' +
          '        <!-- Order icon -->' +
          '        <bt:Image id="order-16" DefaultValue="' + AddinManifestManagement.XMLEncode(WebClientUrl()) + '/Resources/Images/Order_16x16.png"/>' +
          '        <bt:Image id="order-32" DefaultValue="' + AddinManifestManagement.XMLEncode(WebClientUrl()) + '/Resources/Images/Order_32x32.png"/>' +
          '        <bt:Image id="order-80" DefaultValue="' + AddinManifestManagement.XMLEncode(WebClientUrl()) + '/Resources/Images/Order_80x80.png"/>' +
          '      </bt:Images>' +
          '      <bt:Urls>' +
          '        <bt:Url id="taskPaneUrl" DefaultValue="' + AddinManifestManagement.XMLEncode(addinManifestManagement.ConstructURL('OutlookTaskPane', 'taskPaneUrl', GetManifestVersion())) + '"/>' +
          '        <bt:Url id="ShowTaskpane" DefaultValue="' + AddinManifestManagement.XMLEncode(addinManifestManagement.ConstructURL('OutlookTaskPane', 'ShowTaskpane', GetManifestVersion())) + '"/>';
        CustomOutlookNewAction.Reset;
        if CustomOutlookNewAction.FindSet then
            repeat
                Value := Value + CustomOutlookNewAction.GetUrlNode;
            until CustomOutlookNewAction.Next = 0;

        Value := Value + '</bt:Urls>' +
          '      <bt:ShortStrings>' +
          '        <!-- Both modes -->' +
          '        <bt:String id="groupLabel" DefaultValue="' + AddinManifestManagement.XMLEncode(PRODUCTNAME.Short()) + '"/>' +
          '' +
          '        <bt:String id="openPaneButtonLabel" DefaultValue="Contact Insights"/>' +
          '        <bt:String id="openPaneSuperTipTitle" DefaultValue="Open ' +
          AddinManifestManagement.XMLEncode(PRODUCTNAME.Short()) + ' in Outlook"/>' +
          '' +
          '        <bt:String id="newMenuButtonLabel" DefaultValue="New"/>' +
          '        <bt:String id="newMenuSuperTipTitle" DefaultValue="Create a new document in ' +
          AddinManifestManagement.XMLEncode(PRODUCTNAME.Short()) + '"/>' +
          '' +
          '        <bt:String id="pageMenuButtonLabel" DefaultValue="Page"/>' +
          '        <bt:String id="pageMenuSuperTipTitle" DefaultValue="Opens a page in ' +
          AddinManifestManagement.XMLEncode(PRODUCTNAME.Short()) + '"/>' +
          '';

        CustomOutlookNewAction.Reset;
        if CustomOutlookNewAction.FindSet then
            repeat
                Value := Value + CustomOutlookNewAction.GetLabelNode;
                Value := Value + CustomOutlookNewAction.GetSuperTipTitleNode;
            until CustomOutlookNewAction.Next = 0;

        Value := Value + '</bt:ShortStrings>' +
          '      <bt:LongStrings>' +
          '        <bt:String id="groupTooltip" DefaultValue="' + AddinManifestManagement.XMLEncode(PRODUCTNAME.Short()) + ' Add-in"/>' +
          '' +
          '        <bt:String id="openPaneButtonTooltip" DefaultValue="Opens the contact in an embedded view"/>' +
          '        <bt:String id="openPaneSuperTipDesc" DefaultValue="Opens a pane to interact with the customer or vendor"/>' +
          '' +
          '        <bt:String id="newMenuButtonTooltip" DefaultValue="Creates a new document in ' +
          AddinManifestManagement.XMLEncode(PRODUCTNAME.Short()) + '"/>' +
          '        <bt:String id="newMenuSuperTipDesc" DefaultValue="Creates a new document for the selected customer or vendor"/>' +
          '' +
          '        <bt:String id="pageMenuButtonTooltip" DefaultValue="Opens a page in ' +
          AddinManifestManagement.XMLEncode(PRODUCTNAME.Short) + '"/>' +
          '        <bt:String id="pageMenuSuperTipDesc" DefaultValue="Opens a page in ' +
          AddinManifestManagement.XMLEncode(PRODUCTNAME.Short) + '"/>' +
          '';
        CustomOutlookNewAction.Reset;
        if CustomOutlookNewAction.FindSet then
            repeat
                Value := Value + CustomOutlookNewAction.GetTipNode;
                Value := Value + CustomOutlookNewAction.GetSuperTipDescNode;
            until CustomOutlookNewAction.Next = 0;

        Value := Value + '</bt:LongStrings>' +
          '    </Resources>' +
          '  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1"' +
          ' xsi:type="VersionOverridesV1_1">' +
          '    <Requirements>' +
          '      <bt:Sets DefaultMinVersion="1.5">' +
          '        <bt:Set Name="Mailbox" />' +
          '      </bt:Sets>' +
          '    </Requirements>' +
          '    <Hosts>' +
          '      <Host xsi:type="MailHost">' +
          '        <DesktopFormFactor>' +
          '          <!-- Custom pane, only applies to read form -->' +
          '          <ExtensionPoint xsi:type="CustomPane">' +
          '            <RequestedHeight>300</RequestedHeight>' +
          '            <SourceLocation resid="taskPaneUrl"/>' +
          '            <!-- Change this Mode to Or to enable the custom pane -->' +
          '            <Rule xsi:type="RuleCollection" Mode="And">' +
          '              <Rule xsi:type="ItemIs" ItemType="Message"/>' +
          '              <Rule xsi:type="ItemIs" ItemType="AppointmentAttendee"/>' +
          '            </Rule>' +
          '          </ExtensionPoint>' +
          '' +
          '          <!-- Message read form -->' +
          '          <ExtensionPoint xsi:type="MessageReadCommandSurface">' +
          '            <OfficeTab id="TabDefault">' +
          '              <Group id="msgReadGroup">' +
          '                <Label resid="groupLabel" />' +
          '                <Tooltip resid="groupTooltip" />' +
          '' +
          '                <!-- Task pane button -->' +
          '                <Control xsi:type="Button" id="msgReadOpenPaneButton">' +
          '                  <Label resid="openPaneButtonLabel" />' +
          '                  <Tooltip resid="openPaneButtonTooltip" />' +
          '                  <Supertip>' +
          '                    <Title resid="openPaneSuperTipTitle" />' +
          '                    <Description resid="openPaneSuperTipDesc" />' +
          '                  </Supertip>' +
          '                  <Icon>' +
          '                    <bt:Image size="16" resid="nav-icon-16" />' +
          '                    <bt:Image size="32" resid="nav-icon-32" />' +
          '                    <bt:Image size="80" resid="nav-icon-80" />' +
          '                  </Icon>' +
          '                  <Action xsi:type="ShowTaskpane">' +
          '                    <SourceLocation resid="taskPaneUrl" />' +
          '                    <SupportsPinning>true</SupportsPinning>' +
          '                  </Action>' +
          '                </Control>' +
          '';
        CustomOutlookNewAction.Reset;
        CustomOutlookNewAction.SetFilter("Table ID", '<>%1', 0);
        if CustomOutlookNewAction.FindSet then begin
            Value := Value +
        '                <!-- New Page group -->' +
        '                <Control xsi:type="Menu" id="newMenuReadButton">' +
        '                  <Label resid="newMenuButtonLabel" />' +
        '                  <Tooltip resid="newMenuButtonTooltip" />' +
        '                  <Supertip>' +
        '                    <Title resid="newMenuSuperTipTitle" />' +
        '                    <Description resid="newMenuSuperTipDesc" />' +
        '                  </Supertip>' +
        '                  <Icon>' +
        '                    <bt:Image size="16" resid="new-document-16" />' +
        '                    <bt:Image size="32" resid="new-document-32" />' +
        '                    <bt:Image size="80" resid="new-document-80" />' +
        '                  </Icon>' +
        '                  <Items>';
            i := 1;
            CustomOutlookNewAction.Reset;
            CustomOutlookNewAction.SetFilter("Table ID", '<>%1', 0);
            if CustomOutlookNewAction.FindSet then
                repeat
                    Value := Value + CustomOutlookNewAction.GetMenuReadItem(i, 'ReadOpenNew');
                    i := i + 1;
                until CustomOutlookNewAction.Next = 0;

            Value := Value +
            '                  </Items>' +
            '                </Control>';
        end;
        CustomOutlookNewAction.Reset;
        CustomOutlookNewAction.SetRange("Table ID", 0);
        if CustomOutlookNewAction.FindSet then begin
            Value := Value +
            '                <!-- 4PS Page group -->' +
            '                <Control xsi:type="Menu" id="pageMenuReadButton">' +
            '                  <Label resid="pageMenuButtonLabel" />' +
            '                  <Tooltip resid="pageMenuButtonTooltip" />' +
            '                  <Supertip>' +
            '                    <Title resid="pageMenuSuperTipTitle" />' +
            '                    <Description resid="pageMenuSuperTipDesc" />' +
            '                  </Supertip>' +
            '                  <Icon>' +
            '                    <bt:Image size="16" resid="nav-icon-16" />' +
            '                    <bt:Image size="32" resid="nav-icon-32" />' +
            '                    <bt:Image size="80" resid="nav-icon-80" />' +
            '                  </Icon>' +
            '                  <Items>';
            i := 1;
            CustomOutlookNewAction.Reset;
            CustomOutlookNewAction.SetRange("Table ID", 0);
            if CustomOutlookNewAction.FindSet then
                repeat
                    Value := Value + CustomOutlookNewAction.GetMenuReadItem(i, 'ReadOpenOpen');
                    i := i + 1;
                until CustomOutlookNewAction.Next = 0;


            Value := Value + '  </Items>' +
            '                </Control>';
        end;
        Value := Value +
          '              </Group>' +
          '            </OfficeTab>' +
          '          </ExtensionPoint>' +
          '' +
          '          <!-- Message compose form -->' +
          '          <ExtensionPoint xsi:type="MessageComposeCommandSurface">' +
          '            <OfficeTab id="TabDefault">' +
          '              <Group id="msgComposeGroup">' +
          '                <Label resid="groupLabel" />' +
          '                <Tooltip resid="groupTooltip" />' +
          '' +
          '                <!-- Task pane button -->' +
          '                <Control xsi:type="Button" id="msgComposeOpenPaneButton">' +
          '                  <Label resid="openPaneButtonLabel" />' +
          '                  <Tooltip resid="openPaneButtonTooltip" />' +
          '                  <Supertip>' +
          '                    <Title resid="openPaneSuperTipTitle" />' +
          '                    <Description resid="openPaneSuperTipDesc" />' +
          '                  </Supertip>' +
          '                  <Icon>' +
          '                    <bt:Image size="16" resid="nav-icon-16" />' +
          '                    <bt:Image size="32" resid="nav-icon-32" />' +
          '                    <bt:Image size="80" resid="nav-icon-80" />' +
          '                  </Icon>' +
          '                  <Action xsi:type="ShowTaskpane">' +
          '                    <SourceLocation resid="taskPaneUrl" />' +
          '                  </Action>' +
          '                </Control>' +
          '';
        CustomOutlookNewAction.Reset;
        CustomOutlookNewAction.SetFilter("Table ID", '<>%1', 0);
        if CustomOutlookNewAction.FindSet then begin
            Value := Value +
        '                <!-- New Page group -->' +
        '                <Control xsi:type="Menu" id="newMenuComposeButton">' +
        '                  <Label resid="newMenuButtonLabel" />' +
        '                  <Tooltip resid="newMenuButtonTooltip" />' +
        '                  <Supertip>' +
        '                    <Title resid="newMenuSuperTipTitle" />' +
        '                    <Description resid="newMenuSuperTipDesc" />' +
        '                  </Supertip>' +
        '                  <Icon>' +
        '                    <bt:Image size="16" resid="new-document-16" />' +
        '                    <bt:Image size="32" resid="new-document-32" />' +
        '                    <bt:Image size="80" resid="new-document-80" />' +
        '                  </Icon>' +
        '                  <Items>';
            i := 1;
            CustomOutlookNewAction.Reset;
            CustomOutlookNewAction.SetFilter("Table ID", '<>%1', 0);
            if CustomOutlookNewAction.FindSet then
                repeat
                    Value := Value + CustomOutlookNewAction.GetMenuReadItem(i, 'ComposeOpenNew');
                    i := i + 1;
                until CustomOutlookNewAction.Next = 0;

            Value := Value +
            '                  </Items>' +
            '                </Control>';
        end;
        CustomOutlookNewAction.Reset;
        CustomOutlookNewAction.SetRange("Table ID", 0);
        if CustomOutlookNewAction.FindSet then begin
            Value := Value +
            '                <!-- 4PS Page group -->' +
            '                <Control xsi:type="Menu" id="pageMenuComposeButton">' +
            '                  <Label resid="pageMenuButtonLabel" />' +
            '                  <Tooltip resid="pageMenuButtonTooltip" />' +
            '                  <Supertip>' +
            '                    <Title resid="pageMenuSuperTipTitle" />' +
            '                    <Description resid="pageMenuSuperTipDesc" />' +
            '                  </Supertip>' +
            '                  <Icon>' +
            '                    <bt:Image size="16" resid="nav-icon-16" />' +
            '                    <bt:Image size="32" resid="nav-icon-32" />' +
            '                    <bt:Image size="80" resid="nav-icon-80" />' +
            '                  </Icon>' +
            '                  <Items>';
            i := 1;
            CustomOutlookNewAction.Reset;
            CustomOutlookNewAction.SetRange("Table ID", 0);
            if CustomOutlookNewAction.FindSet then
                repeat
                    Value := Value + CustomOutlookNewAction.GetMenuReadItem(i, 'ComposeOpenOpen');
                    i := i + 1;
                until CustomOutlookNewAction.Next = 0;


            Value := Value + '  </Items>' +
             '                </Control>';
        end;

        Value := Value +
          '              </Group > ' +
          '            </OfficeTab>' +
          '          </ExtensionPoint>' +
          '' +
          '          <!-- Appointment organizer form -->' +
          '          <ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">' +
          '            <OfficeTab id="TabDefault">' +
          '              <Group id="apptOrganizerGroup">' +
          '                <Label resid="groupLabel" />' +
          '                <Tooltip resid="groupTooltip" />' +
          '                <!-- Task pane button -->' +
          '                <Control xsi:type="Button" id="apptOrganizerOpenPaneButton">' +
          '                  <Label resid="openPaneButtonLabel" />' +
          '                  <Tooltip resid="openPaneButtonTooltip" />' +
          '                  <Supertip>' +
          '                    <Title resid="openPaneSuperTipTitle" />' +
          '                    <Description resid="openPaneSuperTipDesc" />' +
          '                  </Supertip>' +
          '                  <Icon>' +
          '                    <bt:Image size="16" resid="nav-icon-16" />' +
          '                    <bt:Image size="32" resid="nav-icon-32" />' +
          '                    <bt:Image size="80" resid="nav-icon-80" />' +
          '                  </Icon>' +
          '                  <Action xsi:type="ShowTaskpane">' +
          '                    <SourceLocation resid="taskPaneUrl" />' +
          '                  </Action>' +
          '                </Control>';
        CustomOutlookNewAction.Reset;
        CustomOutlookNewAction.SetFilter("Table ID", '<>%1', 0);
        if CustomOutlookNewAction.FindSet then begin
            Value := Value +
        '                <!-- New Page group -->' +
        '                <Control xsi:type="Menu" id="newMenuOrganizerButton">' +
        '                  <Label resid="newMenuButtonLabel" />' +
        '                  <Tooltip resid="newMenuButtonTooltip" />' +
        '                  <Supertip>' +
        '                    <Title resid="newMenuSuperTipTitle" />' +
        '                    <Description resid="newMenuSuperTipDesc" />' +
        '                  </Supertip>' +
        '                  <Icon>' +
        '                    <bt:Image size="16" resid="new-document-16" />' +
        '                    <bt:Image size="32" resid="new-document-32" />' +
        '                    <bt:Image size="80" resid="new-document-80" />' +
        '                  </Icon>' +
        '                  <Items>';
            i := 1;
            CustomOutlookNewAction.Reset;
            CustomOutlookNewAction.SetFilter("Table ID", '<>%1', 0);
            if CustomOutlookNewAction.FindSet then
                repeat
                    Value := Value + CustomOutlookNewAction.GetMenuReadItem(i, 'OrganizerNew');
                    i := i + 1;
                until CustomOutlookNewAction.Next = 0;

            Value := Value +
            '                  </Items>' +
            '                </Control>';
        end;
        CustomOutlookNewAction.Reset;
        CustomOutlookNewAction.SetRange("Table ID", 0);
        if CustomOutlookNewAction.FindSet then begin
            Value := Value +
            '                <!-- 4PS Page group -->' +
            '                <Control xsi:type="Menu" id="pageMenuOrganizerButton">' +
            '                  <Label resid="pageMenuButtonLabel" />' +
            '                  <Tooltip resid="pageMenuButtonTooltip" />' +
            '                  <Supertip>' +
            '                    <Title resid="pageMenuSuperTipTitle" />' +
            '                    <Description resid="pageMenuSuperTipDesc" />' +
            '                  </Supertip>' +
            '                  <Icon>' +
            '                    <bt:Image size="16" resid="nav-icon-16" />' +
            '                    <bt:Image size="32" resid="nav-icon-32" />' +
            '                    <bt:Image size="80" resid="nav-icon-80" />' +
            '                  </Icon>' +
            '                  <Items>';
            i := 1;
            CustomOutlookNewAction.Reset;
            CustomOutlookNewAction.SetRange("Table ID", 0);
            if CustomOutlookNewAction.FindSet then
                repeat
                    Value := Value + CustomOutlookNewAction.GetMenuReadItem(i, 'OrganizerOpen');
                    i := i + 1;
                until CustomOutlookNewAction.Next = 0;


            Value := Value + '  </Items>' +
            '                </Control>';
        end;
        Value := Value +
          '              </Group>' +
          '            </OfficeTab>' +
          '          </ExtensionPoint>' +
          '' +
          '          <!-- Appointment attendee form -->' +
          '          <ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">' +
          '            <OfficeTab id="TabDefault">' +
          '              <Group id="apptAttendeeGroup">' +
          '                <Label resid="groupLabel" />' +
          '                <Tooltip resid="groupTooltip" />' +
          '                <!-- Task pane button -->' +
          '                <Control xsi:type="Button" id="apptAttendeeOpenPaneButton">' +
          '                  <Label resid="openPaneButtonLabel" />' +
          '                  <Tooltip resid="openPaneButtonTooltip" />' +
          '                  <Supertip>' +
          '                    <Title resid="openPaneSuperTipTitle" />' +
          '                    <Description resid="openPaneSuperTipDesc" />' +
          '                  </Supertip>' +
          '                  <Icon>' +
          '                    <bt:Image size="16" resid="nav-icon-16" />' +
          '                    <bt:Image size="32" resid="nav-icon-32" />' +
          '                    <bt:Image size="80" resid="nav-icon-80" />' +
          '                  </Icon>' +
          '                  <Action xsi:type="ShowTaskpane">' +
          '                    <SourceLocation resid="taskPaneUrl" />' +
          '                  </Action>' +
          '                </Control>' +
          '                <!-- Invoice (dropdown) button -->';
        CustomOutlookNewAction.Reset;
        CustomOutlookNewAction.SetFilter("Table ID", '<>%1', 0);
        if CustomOutlookNewAction.FindSet then begin
            Value := Value +
        '                <!-- New Page group -->' +
        '                <Control xsi:type="Menu" id="newMenuAttendeeButton">' +
        '                  <Label resid="newMenuButtonLabel" />' +
        '                  <Tooltip resid="newMenuButtonTooltip" />' +
        '                  <Supertip>' +
        '                    <Title resid="newMenuSuperTipTitle" />' +
        '                    <Description resid="newMenuSuperTipDesc" />' +
        '                  </Supertip>' +
        '                  <Icon>' +
        '                    <bt:Image size="16" resid="new-document-16" />' +
        '                    <bt:Image size="32" resid="new-document-32" />' +
        '                    <bt:Image size="80" resid="new-document-80" />' +
        '                  </Icon>' +
        '                  <Items>';
            i := 1;
            CustomOutlookNewAction.Reset;
            CustomOutlookNewAction.SetFilter("Table ID", '<>%1', 0);
            if CustomOutlookNewAction.FindSet then
                repeat
                    Value := Value + CustomOutlookNewAction.GetMenuReadItem(i, 'AttendeeNew');
                    i := i + 1;
                until CustomOutlookNewAction.Next = 0;

            Value := Value +
            '                  </Items>' +
            '                </Control>';
        end;
        CustomOutlookNewAction.Reset;
        CustomOutlookNewAction.SetRange("Table ID", 0);
        if CustomOutlookNewAction.FindSet then begin
            Value := Value +
            '                <!-- 4PS Page group -->' +
            '                <Control xsi:type="Menu" id="pageMenuAttendeeButton">' +
            '                  <Label resid="pageMenuButtonLabel" />' +
            '                  <Tooltip resid="pageMenuButtonTooltip" />' +
            '                  <Supertip>' +
            '                    <Title resid="pageMenuSuperTipTitle" />' +
            '                    <Description resid="pageMenuSuperTipDesc" />' +
            '                  </Supertip>' +
            '                  <Icon>' +
            '                    <bt:Image size="16" resid="nav-icon-16" />' +
            '                    <bt:Image size="32" resid="nav-icon-32" />' +
            '                    <bt:Image size="80" resid="nav-icon-80" />' +
            '                  </Icon>' +
            '                  <Items>';
            i := 1;
            CustomOutlookNewAction.Reset;
            CustomOutlookNewAction.SetRange("Table ID", 0);
            if CustomOutlookNewAction.FindSet then
                repeat
                    Value := Value + CustomOutlookNewAction.GetMenuReadItem(i, 'AttendeeOpen');
                    i := i + 1;
                until CustomOutlookNewAction.Next = 0;


            Value := Value + '  </Items>' +
            '                </Control>';
        end;
        Value := Value +
          '              </Group>' +
          '            </OfficeTab>' +
          '          </ExtensionPoint>' +
          '        </DesktopFormFactor>' +
          '      </Host>' +
          '    </Hosts>' +
          '    <Resources>' +
          '      <bt:Images>' +
          '        <!-- NAV icon -->' +
          '        <bt:Image id="nav-icon-16" DefaultValue="' + AddinManifestManagement.XMLEncode(WebClientUrl()) + '/Resources/Images/';
        if EnvironmentInfo.IsSaaS() then
            Value := Value + BrandingFolderTxt + '/OfficeAddin_16x.png"/>'
        else
            Value := Value + '/OfficeAddin_16x.png"/>';

        Value := Value +
          '        <bt:Image id="nav-icon-32" DefaultValue="' + AddinManifestManagement.XMLEncode(WebClientUrl()) + '/Resources/Images/';
        if EnvironmentInfo.IsSaaS() then
            Value := Value + BrandingFolderTxt + '/OfficeAddin_32x.png"/>'
        else
            Value := Value + '/OfficeAddin_32x.png"/>';

        Value := Value +
          '        <bt:Image id="nav-icon-80" DefaultValue="' + AddinManifestManagement.XMLEncode(WebClientUrl()) + '/Resources/Images/';
        if EnvironmentInfo.IsSaaS() then
            Value := Value + BrandingFolderTxt + '/OfficeAddin_80x.png"/>'
        else
            Value := Value + '/OfficeAddin_80x.png"/>';

        Value := Value +
          '' +
          '        <!-- New document icon -->' +
          '        <bt:Image id="new-document-16" DefaultValue="' + AddinManifestManagement.XMLEncode(WebClientUrl()) + '/Resources/Images/NewDocument_16x16.png"/>' +
          '        <bt:Image id="new-document-32" DefaultValue="' + AddinManifestManagement.XMLEncode(WebClientUrl()) + '/Resources/Images/NewDocument_32x32.png"/>' +
          '        <bt:Image id="new-document-80" DefaultValue="' + AddinManifestManagement.XMLEncode(WebClientUrl()) + '/Resources/Images/NewDocument_80x80.png"/>' +
          '        <!-- Order icon -->' +
          '        <bt:Image id="order-16" DefaultValue="' + AddinManifestManagement.XMLEncode(WebClientUrl()) + '/Resources/Images/Order_16x16.png"/>' +
          '        <bt:Image id="order-32" DefaultValue="' + AddinManifestManagement.XMLEncode(WebClientUrl()) + '/Resources/Images/Order_32x32.png"/>' +
          '        <bt:Image id="order-80" DefaultValue="' + AddinManifestManagement.XMLEncode(WebClientUrl()) + '/Resources/Images/Order_80x80.png"/>' +
          '      </bt:Images>' +
          '      <bt:Urls>' +
          '        <bt:Url id="taskPaneUrl" DefaultValue="' + AddinManifestManagement.XMLEncode(addinManifestManagement.ConstructURL('OutlookTaskPane', 'taskPaneUrl', GetManifestVersion())) + '"/>' +
          '        <bt:Url id="ShowTaskpane" DefaultValue="' + AddinManifestManagement.XMLEncode(addinManifestManagement.ConstructURL('OutlookTaskPane', 'ShowTaskpane', GetManifestVersion())) + '"/>';
        CustomOutlookNewAction.Reset;
        if CustomOutlookNewAction.FindSet then
            repeat
                Value := Value + CustomOutlookNewAction.GetUrlNode;
            until CustomOutlookNewAction.Next = 0;

        Value := Value + '</bt:Urls>' +
          '      <bt:ShortStrings>' +
          '        <!-- Both modes -->' +
          '        <bt:String id="groupLabel" DefaultValue="' + AddinManifestManagement.XMLEncode(PRODUCTNAME.Short()) + '"/>' +
          '' +
          '        <bt:String id="openPaneButtonLabel" DefaultValue="Contact Insights"/>' +
          '        <bt:String id="openPaneSuperTipTitle" DefaultValue="Open ' +
          AddinManifestManagement.XMLEncode(PRODUCTNAME.Short()) + ' in Outlook"/>' +
          '' +
          '        <bt:String id="newMenuButtonLabel" DefaultValue="New"/>' +
          '        <bt:String id="newMenuSuperTipTitle" DefaultValue="Create a new document in ' +
          AddinManifestManagement.XMLEncode(PRODUCTNAME.Short()) + '"/>' +
          '' +
          '        <bt:String id="pageMenuButtonLabel" DefaultValue="Page"/>' +
          '        <bt:String id="pageMenuSuperTipTitle" DefaultValue="Opens a page in ' +
          AddinManifestManagement.XMLEncode(PRODUCTNAME.Short()) + '"/>' +
          '';

        CustomOutlookNewAction.Reset;
        if CustomOutlookNewAction.FindSet then
            repeat
                Value := Value + CustomOutlookNewAction.GetLabelNode;
                Value := Value + CustomOutlookNewAction.GetSuperTipTitleNode;
            until CustomOutlookNewAction.Next = 0;

        Value := Value + '      </bt:ShortStrings>' +
       '      <bt:LongStrings>' +
       '        <bt:String id="groupTooltip" DefaultValue="' + AddinManifestManagement.XMLEncode(PRODUCTNAME.Short()) + ' Add-in"/>' +
       '' +
       '        <bt:String id="openPaneButtonTooltip" DefaultValue="Opens the contact in an embedded view"/>' +
       '        <bt:String id="openPaneSuperTipDesc" DefaultValue="Opens a pane to interact with the customer or vendor"/>' +
       '' +
       '        <bt:String id="newMenuButtonTooltip" DefaultValue="Creates a new document in ' +
       AddinManifestManagement.XMLEncode(PRODUCTNAME.Short()) + '"/>' +
       '        <bt:String id="newMenuSuperTipDesc" DefaultValue="Creates a new document for the selected customer or vendor"/>' +
       '' +
       '        <bt:String id="pageMenuButtonTooltip" DefaultValue="Opens a page in ' +
       AddinManifestManagement.XMLEncode(PRODUCTNAME.Short) + '"/>' +
       '        <bt:String id="pageMenuSuperTipDesc" DefaultValue="Opens a page in ' +
       AddinManifestManagement.XMLEncode(PRODUCTNAME.Short) + '"/>' +
       '';
        CustomOutlookNewAction.Reset;
        if CustomOutlookNewAction.FindSet then
            repeat
                Value := Value + CustomOutlookNewAction.GetTipNode;
                Value := Value + CustomOutlookNewAction.GetSuperTipDescNode;
            until CustomOutlookNewAction.Next = 0;

        Value := Value + '</bt:LongStrings>' +
       '    </Resources>' +
       '  </VersionOverrides>' +
       '  </VersionOverrides>' +
       '</OfficeApp>';
    end;
}

