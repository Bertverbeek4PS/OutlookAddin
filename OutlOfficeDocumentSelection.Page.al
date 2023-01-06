page 50109 "Outl Office Doc Selection"
{
    Caption = 'Document Selection';
    DeleteAllowed = false;
    Editable = false;
    InsertAllowed = false;
    ModifyAllowed = false;
    PageType = List;
    SourceTable = "Office Document Selection";
    SourceTableTemporary = true;

    layout
    {
        area(content)
        {
            repeater(Control2)
            {
                ShowCaption = false;
                field(DocumentType; DocumentType)
                {
                    ToolTip = 'Specifies the document type that the entry belongs to.';
                }
                field("Document No."; Rec."Document No.")
                {
                    Lookup = true;
                    ToolTip = 'Specifies the number of the involved document.';
                }
                field(RecorId; Rec."Record ID")
                {
                    Lookup = true;
                    ToolTip = 'Specifies the number of the involved document.';
                }
            }
        }
    }

    actions
    {
        area(navigation)
        {
            action("View Document")
            {
                ApplicationArea = Basic, Suite;
                Caption = 'View Document';
                Image = ViewOrder;
                ShortCutKey = 'Return';
                ToolTip = 'View the selected document.';

                trigger OnAction()
                var
                    HyperlinkgManifest4PS: Codeunit "Hyperlink Manifest BV";
                begin
                    HyperlinkgManifest4PS.OpenIndividualDocument(Rec);
                end;
            }
        }
        area(Promoted)
        {
            group(Category_Process)
            {
                Caption = 'Process';

                actionref("View Document_Promoted"; "View Document")
                {
                }
            }
        }
    }
    var
        DocumentType: Text;

    trigger OnAfterGetRecord()
    var
        RecRef: RecordRef;
    begin
        RecRef.Close();
        RecRef.Open(rec."Table ID");
        DocumentType := RecRef.Caption;
    end;

}