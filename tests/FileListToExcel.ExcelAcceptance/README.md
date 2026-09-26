# Developer-only live Excel acceptance harness

This tool launches a separate normal Excel Desktop process with /x and a dedicated synthetic workbook. It resolves that exact process through Excel's native object model, observes COMAddIns and menus, and applies test selections/filters/sorts. It never sets Connect, EnableEvents, AutomationSecurity, VBA trust or policy. It is not an end-user setup or runtime dependency.

Build on Windows with the repository SDK. Invoke the built executable with two arguments: the absolute synthetic .xlsx fixture and a new control/evidence directory. The current launch path assumes Office Click-to-Run under Program Files/Microsoft Office/root/Office16. It writes started.json; subsequent command.json objects have a strictly increasing id and an operation, with results in result-ID.json. Consult Program.cs for supported operations. Normal close dismisses this instance's test workbooks without saving and calls Quit; never use user documents or mix user work into this test instance.

The execute operation invokes the real compiled CommandBar click callback; destination choice and actual GUI inspection remain interactive acceptance steps. No selection is read from a saved workbook in place of Excel's live state. A menu snapshot immediately after a COM command alone is not proof of correct displayed UI; inspect the current window and use the same installed MSI for the final acceptance run.

Run in a disposable profile/VM for release approval. This repository's recorded live run used the existing host and separately lists unexecuted environments. Evidence and local paths under artifacts must not be published without review.
