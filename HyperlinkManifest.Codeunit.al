codeunit 50104 "Hyperlink Manifest BV"
{
    TableNo = "Office Add-in Context";

    trigger OnRun()
    begin
        RedirectToDocument(Rec);
    end;

    var
        SalesReceivablesSetup: Record "Sales & Receivables Setup";
        PurchasePayablesSetup: Record "Purchases & Payables Setup";
        ServiceSetup: Record "Service Setup";
        AddInManifestManagement: Codeunit "Add-in Manifest Management";
        AddinNameTxt: Label 'Document View 4PS';
        AddinDescriptionTxt: Label 'Provides a link directly to business documents in %1.', Comment = '%1 - Application Name';
        AppIdTxt: Label '6e8ffb1d-5dc3-4546-8bd7-b419dd6f8f30', Locked = true;
        PurchaseOrderAcronymTxt: Label 'PO', Comment = 'US acronym for Purchase Order';
        DocDoesNotExistMsg: Label 'Cannot find a document with the number %1.', Comment = '%1=The document number the hyperlink is attempting to open.';
        SuggestedItemsDisabledTxt: Label 'The suggested line items page has been disabled by the user.', Locked = true;
        DocumentMatchedTelemetryTxt: Label 'Outlook Document View loaded%1  Documents matched: %2%1  Document Series: %3%1  Document Type: %4', Locked = true;
        BrandingFolderTxt: Label 'ProjectMadeira', Locked = true;
        EnvironmentInfo: Codeunit "Environment Information";

    internal procedure RedirectToDocument(TempOfficeAddinContext: Record "Office Add-in Context" temporary)
    var
        TempOfficeDocumentSelection: Record "Office Document Selection" temporary;
        TypeHelper: Codeunit "Type Helper";
        DocNo: Code[20];
    begin
        DocNo := CopyStr(TempOfficeAddinContext."Regular Expression Match", 1, 20);
        CollectDocumentMatches(TempOfficeDocumentSelection, DocNo, TempOfficeAddinContext);

        Session.LogMessage('0000ACS', StrSubstNo(DocumentMatchedTelemetryTxt,
                TypeHelper.NewLine(),
                TempOfficeDocumentSelection.Count(),
                Format(TempOfficeDocumentSelection.Series),
                Format(TempOfficeDocumentSelection."Document Type")), Verbosity::Normal, DataClassification::SystemMetadata, TelemetryScope::ExtensionPublisher, 'Category', 'AL Office Add-in');

        case TempOfficeDocumentSelection.Count of
            0:
                begin
                    TempOfficeAddinContext."Document No." := DocNo;
                    PAGE.Run(PAGE::"Office Doc Selection Dlg");
                end;
            1:
                OpenIndividualDocument(TempOfficeDocumentSelection);
            else // More than one document match, must have user pick
                PAGE.Run(PAGE::"Outl Office Doc Selection", TempOfficeDocumentSelection);
        end;
    end;

    local procedure CollectDocumentMatches(var TempOfficeDocumentSelection: Record "Office Document Selection" temporary; var DocNo: Code[20]; TempOfficeAddinContext: Record "Office Add-in Context" temporary)
    begin
        if (TempOfficeAddinContext."Document No." = '') and (TempOfficeAddinContext."Regular Expression Match" <> '') then begin
            //First search for keywords in the expression
            if not ExpressionContainsSeriesTitle(TempOfficeAddinContext."Regular Expression Match", DocNo, TempOfficeDocumentSelection) then begin
                //then search for real document numbers
                SetDocumentMatchRecord(DocNo, TempOfficeDocumentSelection);
            end;
        end;
    end;

    internal procedure OpenIndividualDocument(TempOfficeDocumentSelection: Record "Office Document Selection" temporary)
    var
        RecRef: RecordRef;
        VarRecRef: Variant;
    begin
        RecRef.Close();
        RecRef.Open(TempOfficeDocumentSelection."Table ID");
        if RecRef.Get(TempOfficeDocumentSelection."Record ID") then begin
            VarRecRef := RecRef;
            PAGE.Run(TempOfficeDocumentSelection."Page ID", VarRecRef);
        end else
            message(StrSubstNo(DocDoesNotExistMsg, TempOfficeDocumentSelection."Document No."));
    end;

    local procedure CreateDocumentMatchRecord(var TempOfficeDocumentSelection: Record "Office Document Selection" temporary; Series: Option; DocType: Enum "Incoming Document Type"; DocNo: Code[20]; Posted: Boolean; DocDate: Date; PageId: Integer; TableId: Integer; RecordId: RecordId)
    begin
        TempOfficeDocumentSelection.Reset;
        TempOfficeDocumentSelection.SetRange(Series, Series);
        TempOfficeDocumentSelection.SetRange(Posted, Posted);
        TempOfficeDocumentSelection.SetRange("Document No.", DocNo);
        if TempOfficeDocumentSelection.FindLast() then begin
            TempOfficeDocumentSelection.Validate("Document Type", TempOfficeDocumentSelection."Document Type".AsInteger() + 1);
        end else begin
            TempOfficeDocumentSelection.Validate("Document Type", DocType);
        end;
        TempOfficeDocumentSelection.Init();
        TempOfficeDocumentSelection.Validate("Document No.", DocNo);
        TempOfficeDocumentSelection.Validate("Document Date", DocDate);
        TempOfficeDocumentSelection.Validate(Series, Series);
        TempOfficeDocumentSelection.Validate(Posted, Posted);
        TempOfficeDocumentSelection.Validate("Page ID", PageId);
        TempOfficeDocumentSelection.Validate("Table ID", TableId);
        TempOfficeDocumentSelection.Validate("Record ID", RecordId);
        TempOfficeDocumentSelection.Insert();
    end;

    local procedure ExpressionContainsSeriesTitle(Expression: Text[250]; var DocNo: Code[20]; var TempOfficeDocumentSelection: Record "Office Document Selection" temporary): Boolean
    var
        HyperlinkManifest: Codeunit "Hyperlink Manifest BV";
        OutlDocumentViewer4PS: Record "Outl. Document Viewer";
        Found: Boolean;
    begin
        Found := false;
        if OutlDocumentViewer4PS.FindSet() then
            repeat
                if GetDocumentNumber(Expression, OutlDocumentViewer4PS."Table Caption", DocNo) then begin
                    SetDocumentMatchRecord(DocNo, TempOfficeDocumentSelection);
                    Found := true;
                end;
            until OutlDocumentViewer4PS.Next = 0;

        exit(Found);
    end;

    internal procedure GetDocumentNumber(Expression: Text[250]; Keyword: Text; var DocNo: Code[20]) IsMatch: Boolean
    var
        HyperlinkManifest: Codeunit "Hyperlink Manifest";
        DocNoRegEx: Codeunit Regex;
        Matches: Record Matches;
    begin
        DocNoRegEx.Regex(StrSubstNo('(?i)(?<=%1).*', Keyword));
        DocNoRegEx.Match(Expression, 0, Matches);

        if Matches.FindFirst() then begin
            DocNo := Matches.ReadValue();
            exit(true);
        end;
    end;

    local procedure SetDocumentMatchRecord(DocNo: Code[20]; var TempOfficeDocumentSelection: Record "Office Document Selection" temporary)
    var
        OutlDocumentViewer4PS: Record "Outl. Document Viewer";
        RecRef: RecordRef;
    begin
        if OutlDocumentViewer4PS.FindSet() then
            repeat
                if FindRecordInTable(OutlDocumentViewer4PS, DocNo, RecRef) then
                    CreateDocumentMatchRecord(TempOfficeDocumentSelection,
                                            TempOfficeDocumentSelection.Series::Purchase,
                                            TempOfficeDocumentSelection."Document Type"::Custom,
                                            DocNo,
                                            false,
                                            DT2Date(CurrentDateTime),
                                            OutlDocumentViewer4PS."Page ID",
                                            OutlDocumentViewer4PS."Table ID",
                                            RecRef.RecordId);
            until OutlDocumentViewer4PS.Next = 0;
    end;

    local procedure FindRecordInTable(OutlDocumentViewer4PS: Record "Outl. Document Viewer"; DocNo: Code[20]; var RecRef: RecordRef): Boolean
    var
        PurchaseHeader: Record "Purchase Header";
        SalesHeader: Record "Sales Header";
    begin
        RecRef.Close();
        RecRef.Open(OutlDocumentViewer4PS."Table ID");
        RecRef.Field(OutlDocumentViewer4PS."Field No. Document No.").SetFilter(DocNo);
        //Purchase Header and Sales Header
        Case OutlDocumentViewer4PS."Page ID" of
            41:
                RecRef.Field(1).SetFilter('%1', PurchaseHeader."Document Type"::Quote);
            42:
                RecRef.Field(1).SetFilter('%1', PurchaseHeader."Document Type"::Order);
            43:
                RecRef.Field(1).SetFilter('%1', PurchaseHeader."Document Type"::Invoice);
            44:
                RecRef.Field(1).SetFilter('%1', PurchaseHeader."Document Type"::"Credit Memo");
            49:
                RecRef.Field(1).SetFilter('%1', SalesHeader."Document Type"::Quote);
            50:
                RecRef.Field(1).SetFilter('%1', SalesHeader."Document Type"::Order);
            51:
                RecRef.Field(1).SetFilter('%1', SalesHeader."Document Type"::Invoice);
            52:
                RecRef.Field(1).SetFilter('%1', SalesHeader."Document Type"::"Credit Memo");
        end;

        if RecRef.FindFirst() then
            exit(true)
        else
            exit(false);
    end;

    internal procedure SetHyperlinkAddinTriggers() RegExText: Text
    var
        OutlDocumentViewer4PS: Record "Outl. Document Viewer";
        RegExTextNoSeries: Text;
        RegExTextWords: Text;
    begin
        // First add the number series rules
        OutlDocumentViewer4PS.Reset;
        if OutlDocumentViewer4PS.FindSet() then
            repeat
                RegExTextNoSeries := AddPrefixesToRegex(GetNoSeriesPrefixes(OutlDocumentViewer4PS."No. Series"), RegExTextNoSeries);
            until OutlDocumentViewer4PS.Next = 0;

        // Wrap the prefixes in parenthesis to group them and fill out the rest of the RegEx:
        if RegExTextNoSeries <> '' then begin
            RegExText := StrSubstNo('(%1)([A-Za-z0-9]+)', RegExTextNoSeries);
        end;

        RegExTextWords := 'invoice|order|quote|credit memo'; //Adding some standard words

        OutlDocumentViewer4PS.Reset;
        if OutlDocumentViewer4PS.FindSet() then
            repeat
                RegExTextWords += '|' + OutlDocumentViewer4PS."Table Caption";
            until OutlDocumentViewer4PS.Next = 0;

        if RegExTextNoSeries <> '' then
            RegExText += '|';

        RegExText +=
          StrSubstNo('(%1):? ?#?(%2)', RegExTextWords, GetNumberSeriesRegex());

        exit(RegExText);

    end;

    internal procedure GetPrefixForNoSeriesLine(var NoSeriesLine: Record "No. Series Line"): Code[20]
    var
        NumericRegEx: Codeunit Regex;
        Matches: Record Matches;
        SeriesStartNo: Code[20];
        MatchText: Text;
        LowerMatchBound: Integer;
    begin
        SeriesStartNo := NoSeriesLine."Starting No.";

        // Ensure that we have a non-numeric 'prefix' before the numbers and that we capture the last number group.
        // This ensures that we can generate a specific RegEx and not match all number sequences.
        SeriesStartNo := NumericRegEx.Replace(SeriesStartNo, '[\-\[\]\/\{\}\(\)\*\+\?\.\\\^\$\|]', '');
        NumericRegEx.Regex('^[a-zA-Z]*');
        NumericRegEx.Match(SeriesStartNo, 0, Matches);

        // If we don't have a match, then the code is unusable for a RegEx as a number series
        if Matches.Count = 0 then
            exit('');

        Matches.SetRange(Success, true);
        if Matches.FindFirst() then
            SeriesStartNo := CopyStr(SeriesStartNo, 1, Matches.Length);

        exit(SeriesStartNo);
    end;

    internal procedure GetNoSeriesPrefixes(NoSeriesCode: Code[20]): Text
    var
        NoSeriesLine: Record "No. Series Line";
        NewPrefix: Text;
        Prefixes: Text;
    begin
        // For the given series code - get the prefix for each line
        NoSeriesLine.SetRange("Series Code", NoSeriesCode);
        if NoSeriesLine.Find('-') then
            repeat
                NewPrefix := GetPrefixForNoSeriesLine(NoSeriesLine);
                if NewPrefix <> '' then
                    if Prefixes = '' then
                        Prefixes := RegExEscape(NewPrefix)
                    else
                        Prefixes := StrSubstNo('%1|%2', Prefixes, RegExEscape(NewPrefix));
            until NoSeriesLine.Next() = 0;

        exit(Prefixes);
    end;

    local procedure AddPrefixesToRegex(Prefixes: Text; RegExText: Text): Text
    begin
        // Handles some logic around concatenating the prefixes together in a regex string
        if Prefixes <> '' then
            if RegExText = '' then
                RegExText := Prefixes
            else
                RegExText := StrSubstNo('%1|%2', RegExText, Prefixes);
        exit(RegExText);
    end;

    local procedure GetManifestVersion(): Text
    begin
        exit('2.1.0.0');
    end;

    local procedure RegExEscape(RegExText: Text): Text
    var
        RegEx: Codeunit Regex;
    begin
        // Function to escape some special characters in a regular expression character class:
        exit(RegEx.Escape(RegExText));
    end;

    procedure GetNumberSeriesRegex(): Text
    begin
        exit(StrSubstNo('[\w%1]*[0-9]+', RegExEscape('_/#*+\|-')));
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Add-in Manifest Management", 'CreateDefaultAddins', '', false, false)]
    local procedure OnCreateAddin(var OfficeAddin: Record "Office Add-in")
    begin
        if OfficeAddin.Get(AppIdTxt) then
            OfficeAddin.Delete();

        OfficeAddin.Init();
        OfficeAddin."Application ID" := AppIdTxt;
        OfficeAddin."Manifest Codeunit" := CODEUNIT::"Hyperlink Manifest BV";
        OfficeAddin.Name := AddinNameTxt;
        OfficeAddin.Description := AddinDescriptionTxt;
        OfficeAddin.Version := GetManifestVersion();
        OfficeAddin.Insert(true);

        OfficeAddin.SetDefaultManifestText(DefaultManifestText());
        OfficeAddin.Modify(true);
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Add-in Manifest Management", 'OnGenerateManifest', '', false, false)]
    local procedure OnGenerateManifest(var OfficeAddin: Record "Office Add-in"; var ManifestText: Text; CodeunitID: Integer)
    var
        AddinURL: Text;
    begin
        if not CanHandle(CodeunitID) then
            exit;

        ManifestText := OfficeAddin.GetDefaultManifestText();
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Office Management", 'OnGetExternalHandlerCodeunit', '', true, false)]
    local procedure RegisterExternalHandler(OfficeAddinContext: Record "Office Add-in Context"; HostType: Text; var HandlerCodeunit: Integer)
    var
        CustomOutlookNewAction: Record "Custom Outlook Action";
    begin
        if HostType = 'Outlook-Hyperlink' then
            HandlerCodeunit := CODEUNIT::"Hyperlink Manifest BV";
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Add-in Manifest Management", 'GetAddin', '', false, false)]
    local procedure OnGetAddin(var OfficeAddin: Record "Office Add-in"; CodeunitID: Integer)
    begin
        if CanHandle(CodeunitID) then
            OfficeAddin.Get(AppIdTxt);
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Add-in Manifest Management", 'GetAddinID', '', false, false)]
    local procedure OnGetAddinID(var ID: Text; CodeunitID: Integer)
    begin
        if CanHandle(CodeunitID) then
            ID := AppIdTxt;
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Add-in Manifest Management", 'GetAddinVersion', '', false, false)]
    local procedure OnGetAddinVersion(var Version: Text; CodeunitID: Integer)
    begin
        if CanHandle(CodeunitID) then
            Version := GetManifestVersion();
    end;

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Add-in Manifest Management", 'GetManifestCodeunit', '', false, false)]
    local procedure OnGetCodeunitID(var CodeunitID: Integer; HostType: Text)
    var
    begin
        if HostType in ['Outlook-Hyperlink'] then
            CodeunitID := CODEUNIT::"Hyperlink Manifest BV";
    end;

    local procedure CanHandle(CodeunitID: Integer): Boolean
    begin
        exit(CodeunitID = CODEUNIT::"Hyperlink Manifest BV");
    end;

    local procedure DefaultManifestText() Value: Text
    begin
        Value :=
          '<?xml version="1.0" encoding="utf-8"?>' +
          '<OfficeApp xsi:type="MailApp" ' +
          '     xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
          '     xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"' +
          '     xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"' +
          '     xmlns:o10="http://schemas.microsoft.com/office/mailappversionoverrides"' +
          '     xmlns:o11="http://schemas.microsoft.com/office/mailappversionoverrides/1.1">' +
          '  <Id>' + AppIdTxt + '</Id>' +
          '  <Version>' + GetManifestVersion() + '</Version>' +
          '  <ProviderName>4PS</ProviderName>' +
          '  <DefaultLocale>en-US</DefaultLocale>' +
          '  <DisplayName DefaultValue="' + AddinNameTxt + '" />' +
          '  <Description DefaultValue="' +
          StrSubstNo(AddinDescriptionTxt, AddInManifestManagement.XMLEncode(PRODUCTNAME.Full())) + '" />' +
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
          '      <Set Name="MailBox" MinVersion="1.1" />' +
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
          '  <Rule xsi:type="RuleCollection" Mode="And">' +
          '    <Rule xsi:type="RuleCollection" Mode="Or">' +
          '      <!-- To add more complex rules, add additional rule elements -->' +
          '      <!-- E.g. To activate when a message contains an address -->' +
          '      <!-- <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" /> -->' +
          '      <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="DocumentTypes" RegExValue="' + SetHyperlinkAddinTriggers + '" PropertyName="BodyAsPlaintext" IgnoreCase="true" />' +
          '    </Rule>' +
          '    <Rule xsi:type="RuleCollection" Mode="Or">' +
          '      <Rule xsi:type="ItemIs" FormType="Edit" ItemType="Message" />' +
          '      <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />' +
          '    </Rule>' +
          '' +
          '  </Rule>' +
          '  <DisableEntityHighlighting>false</DisableEntityHighlighting>' +
          '  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">' +
          '    <!-- VersionOverrides for the v1.1 schema -->' +
          '    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">' +
          '      <Requirements>' +
          '        <bt:Sets DefaultMinVersion="1.5">' +
          '          <bt:Set Name="Mailbox" />' +
          '        </bt:Sets>' +
          '      </Requirements>' +
          '      <Hosts>' +
          '        <Host xsi:type="MailHost">' +
          '          <DesktopFormFactor>' +
          '            <!-- DetectedEntity -->' +
          '            <ExtensionPoint xsi:type="DetectedEntity">' +
          '              <Label resid="contextLabel" />' +
          '              <SourceLocation resid="detectedEntityUrl" />' +
          '              <Rule xsi:type="RuleCollection" Mode="And">' +
          '                <Rule xsi:type="ItemIs" ItemType="Message" />' +
          '                <Rule xsi:type="RuleCollection" Mode="Or">' +
          '                  <!-- To add more complex rules, add additional rule elements -->' +
          '                  <!-- E.g. To activate when a message contains an address -->' +
          '                  <!-- <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" /> -->' +
          '                  <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="DocumentTypes" RegExValue="' + SetHyperlinkAddinTriggers + '" PropertyName="BodyAsPlaintext" IgnoreCase="true" />' +
          '                </Rule>' +
          '              </Rule>' +
          '            </ExtensionPoint>' +
          '          </DesktopFormFactor>' +
          '        </Host>' +
          '      </Hosts>' +
          '      <Resources>' +
          '        <bt:Urls>' +
          '          <bt:Url id="detectedEntityUrl" DefaultValue="' + AddinManifestManagement.XMLEncode(addinManifestManagement.ConstructURL('Outlook-Hyperlink', '', GetManifestVersion())) + '"/>' +
          '        </bt:Urls>' +
          '        <bt:ShortStrings>' +
          '          <bt:String id="contextLabel" DefaultValue="' + AddinNameTxt + '"/>' +
          '        </bt:ShortStrings>' +
          '      </Resources>' +
          '    </VersionOverrides>' +
          '  </VersionOverrides>' +
          '</OfficeApp>';
    end;
}