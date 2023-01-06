page 50105 "Outl. Document Viewer"
{
    PageType = List;
    SourceTable = "Outl. Document Viewer";

    layout
    {
        area(Content)
        {
            repeater(GroupName)
            {
                field(TableId; rec."Table ID")
                {
                    TableRelation = AllObjWithCaption."Object ID" where("Object Type" = Const(Table), "Object Subtype" = Const('Normal'));
                    ToolTip = 'Table to look for the record.';
                }
                field(TableCaption; rec."Table Caption")
                {
                    ToolTip = 'Table Caption of the Table ID.';
                }
                field(FielNoDocumentNo; rec."Field No. Document No.")
                {
                    ToolTip = 'Choose the field of the Document No in the table.';

                    trigger OnLookup(var Text: Text): Boolean
                    var
                        Fld: Record "Field";
                        FieldSelection: Codeunit "Field Selection";
                    begin
                        Fld.SetRange(TableNo, Rec."Table ID");
                        if FieldSelection.Open(Fld) then begin
                            Rec.Validate("Field No. Document No.", Fld."No.");
                        end;
                    end;

                }
                field(PageId; rec."Page ID")
                {
                    TableRelation = AllObjWithCaption."Object ID" where("Object Type" = Const(Page));
                    ToolTip = 'The page that will be openend in the Document Viewer add-in.';
                }
                field(PageCaption; rec."Page Caption")
                {
                    ToolTip = 'Page Caption of the Page ID.';
                }
                field(NoSeries; rec."No. Series")
                {
                    TableRelation = "No. Series";
                    ToolTip = 'Choose the No Series to look for in the e-mail.';
                }
            }
        }
    }
}