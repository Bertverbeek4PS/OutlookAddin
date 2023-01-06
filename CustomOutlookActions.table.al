table 50100 "Custom Outlook Action"
{
    Caption = 'Custom Outlook Action';

    fields
    {
        field(10; Name; Text[250])
        {
            Caption = 'Name';

            trigger OnValidate()
            begin
                if Name <> '' then
                    Name := DelChr(Name, '=', ' |!|@|#|$|%|^|&|*|(|)');
            end;
        }
        field(20; Label; Text[250])
        {
            Caption = 'Label';
        }
        field(30; Tip; Text[250])
        {
            Caption = 'Tip';
        }
        field(40; "Super Tip Title"; Text[250])
        {
            Caption = 'Super Tip Title';
        }
        field(50; "Super Tip Desc"; Text[250])
        {
            Caption = 'Super Tip Description';
        }
        field(60; "Page ID"; Integer)
        {
            Caption = 'Page ID';

            trigger OnValidate()
            var
                AllObjWithCaption: Record AllObjWithCaption;
            begin
                if "Page ID" <> 0 then
                    if AllObjWithCaption.Get(AllObjWithCaption."Object Type"::Page, "Page ID") then
                        "Page Caption" := AllObjWithCaption."Object Name"
                    else begin
                        Clear("Page Caption");
                        Error(err002);
                    end;
            end;
        }
        field(70; "Page Caption"; Text[30])
        {
            Caption = 'Page Caption';
            Editable = false;
        }
        field(80; "Table ID"; Integer)
        {
            Caption = 'Table ID';

            trigger OnValidate()
            var
                AllObjWithCaption: Record AllObjWithCaption;
            begin
                if "Page ID" <> 0 then
                    if AllObjWithCaption.Get(AllObjWithCaption."Object Type"::Table, "Table ID") then
                        "Table Caption" := AllObjWithCaption."Object Name"
                    else begin
                        Clear("Table Caption");
                        Error(err003);
                    end;
            end;
        }
        field(90; "Table Caption"; Text[30])
        {
            Caption = 'Table Caption';
            Editable = false;
        }
        field(100; "Contact Type"; Option)
        {
            Caption = 'Contact Type';
            OptionMembers = ,Contact,Customer,Vendor;
        }
        field(110; "Field No. Contact Type"; Integer)
        {
            Caption = 'Field No. Contact Type';
            TableRelation = Field."No." where(TableNo = field("Table ID"));
        }
        field(120; "Field No. Name"; Integer)
        {
            Caption = 'Field No. Name';
            TableRelation = Field."No." where(TableNo = field("Table ID"));
        }
        field(130; "Field No. E-mail"; Integer)
        {
            Caption = 'Field No. E-mail';
            TableRelation = Field."No." where(TableNo = field("Table ID"));
        }
    }

    keys
    {
        key(Key1; Name)
        {
            Clustered = true;
        }
    }

    fieldgroups
    {
    }

    var
        err002: Label 'Invalid page id.';
        err003: Label 'Invalid table id.';

    internal procedure GetMenuReadItem(No: Integer; GroupTxt: Text) Item: Text
    begin
        Item := '<Item id="pageMenu' + GroupTxt + 'Item' + Format(No) + '">';
        Item := Item + '<Label resid="page' + Name + 'Label" />';
        Item := Item + '<Tooltip resid="page' + Name + 'Tip" />';
        Item := Item + '<Supertip>';
        Item := Item + '<Title resid="page' + Name + 'SuperTipTitle" />';
        Item := Item + '<Description resid="page' + Name + 'SuperTipDesc" />';
        Item := Item + '</Supertip>';
        Item := Item + '<Icon><bt:Image size="16" resid="order-16" /><bt:Image size="32" resid="order-32" /><bt:Image size="80" resid="order-80" /></Icon>';
        Item := Item + '<Action xsi:type="ShowTaskpane">';
        Item := Item + '<SourceLocation resid="page' + Name + 'Url" />';
        Item := Item + '</Action>';
        Item := Item + '</Item>';
    end;

    internal procedure GetUrlNode() Node: Text
    var
        AddinManifestManagement: Codeunit "Add-in Manifest Management";
        FPSOfficeHandlerBV: Codeunit "Office Handler BV";
    begin
        Node := '<bt:Url id="page' + Name + 'Url" DefaultValue="' + AddinManifestManagement.XMLEncode(addinManifestManagement.ConstructURL('OutlookTaskPane', GetCommand(), FPSOfficeHandlerBV.GetManifestVersion())) + '"/>';
    end;

    internal procedure GetUrl(): Text
    begin
        exit('page' + Name + 'Url');
    end;

    internal procedure GetCommand(): Text
    begin
        exit('Page-' + Name);
    end;

    internal procedure GetLabelNode() Node: Text
    begin
        Node := '<bt:String id="page' + Name + 'Label" DefaultValue="' + Label + '"/>';
    end;

    internal procedure GetSuperTipTitleNode() Node: Text
    begin
        Node := '<bt:String id="page' + Name + 'SuperTipTitle" DefaultValue="' + "Super Tip Title" + '"/>';
    end;

    internal procedure GetTipNode() Node: Text
    begin
        Node := '<bt:String id="page' + Name + 'Tip" DefaultValue="' + Tip + '" />';
    end;

    internal procedure GetSuperTipDescNode() Node: Text
    begin
        Node := '<bt:String id="page' + Name + 'SuperTipDesc" DefaultValue="' + "Super Tip Desc" + '" />';
    end;
}

