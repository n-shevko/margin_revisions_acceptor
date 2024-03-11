Installation

Video instruction for installation\usage

1. Turn on "Developer tab" on "ribbon command bar"
File -> Options -> Customize -> Customize Ribbon -> Check "Developer" option

2. On "Developer tab" press "Macros" button -> fill in Macro name: AcceptMarginRevisions -> press "Create" button
remove all content and paste the following source code:

Public Sub AcceptMarginRevisions()
  For Each Revision In ActiveDocument.Revisions
    If Revision.Type = wdRevisionDelete Or Revision.FormatDescription <> "" Then
      Revision.Accept
    End If
  Next
End Sub

Usage

1. Go to "Developer tab" on "ribbon command bar"
2. Press "Macros" button, select "AcceptMarginRevisions" macro
3. Press "Run" button