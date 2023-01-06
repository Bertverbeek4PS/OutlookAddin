table 50101 "Outl. Document Viewer"
{
    fields
    {
        field(10; "Table ID"; Integer)
        {
            Caption = 'Table ID';

            trigger OnValidate()
            var
                AllObjWithCaption: Record AllObjWithCaption;
            begin
                if "Table ID" <> 0 then
                    if AllObjWithCaption.Get(AllObjWithCaption."Object Type"::Table, "Table ID") then
                        "Table Caption" := AllObjWithCaption."Object Name"
                    else begin
                        Clear("Table Caption");
                        Error(err003);
                    end;
            end;
        }
        field(20; "Table Caption"; Text[30])
        {
            Caption = 'Table Caption';
            Editable = false;
        }
        field(60; "Field No. Document No."; Integer)
        {
            Caption = 'Field No. Document No.';
            TableRelation = Field."No." where(TableNo = field("Table ID"));
        }
        field(30; "Page ID"; Integer)
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
        field(40; "Page Caption"; Text[30])
        {
            Caption = 'Page Caption';
            Editable = false;
        }
        field(50; "No. Series"; Code[20])
        {
            Caption = 'No. Series';
        }
    }

    keys
    {
        key(Key1; "Table ID", "Page ID")
        {
            Clustered = true;
        }
    }

    var
        err002: Label 'Invalid page id.';
        err003: Label 'Invalid table id.';
}