tableextension 50100 "Outl. Office Doc Selection" extends "Office Document Selection"
{
    fields
    {
        field(11012000; "Page ID"; Integer)
        {
            Caption = 'Page ID';
        }
        field(11012010; "Table ID"; Integer)
        {
            Caption = 'Table ID';
        }
        field(11012020; "Record ID"; RecordId)
        {
            Caption = 'Record ID';
        }
    }
}