# Phishing-Test
Outlook script to detect and signal phishing test mails
This script is written in VBA and is basically a macro in Outlook.
If you need to enable macro's in Outlook, read on. 
If you want to automatically scan new mails for phishing tests, read the whole text.


# Enable Macros in Outlook:

1. **Open Outlook**: Launch Microsoft Outlook on your computer.

2. **Enable Developer Tab**:
   - If you don't already see the "Developer" tab in the Ribbon at the top of the Outlook window, you need to enable it. Here's how:
     - Click on "File" in the top left corner.
     - Click on "Options."
     - In the Outlook Options window, select "Customize Ribbon."
     - On the right side, check the "Developer" option.
     - Click "OK" to save your changes.

**Create a Macro in Outlook:**

1. **Open Visual Basic for Applications (VBA) Editor**:
   - Click on the "Developer" tab in the Ribbon.
   - Click on "Visual Basic" in the "Code" group. This will open the VBA editor.

2. **Insert a Module**:
   - In the VBA editor, you'll see the Project Explorer on the left and a code window on the right.
   - Right-click on "VbaProject (YourOutlookFile)" in the Project Explorer.
   - Select "Insert" > "Module." This will create a new module where you can write your macro code.

3. **Write Your Macro**:
   - In the code window for the module, you can write your macro code using VBA. For example, here's a simple macro that displays a message box when run:
   ```vba
   Sub MyOutlookMacro()
       MsgBox "Hello, Outlook Macro!"
   End Sub
   ```
   
     **At this stage you might want to copy the phishing test detection code from this repository and paste it here.**
    
4. **Save the Macro**:
   - To save your macro, click the "File" menu in the VBA editor and select "Save [YourOutlookFile]."
   - Close the VBA editor.

5. **Run the Macro**:
   - To run the macro, go back to Outlook.
   - Click on the "Developer" tab in the Ribbon.
   - Click on "Macros."
   - Select your macro (e.g., "MyOutlookMacro") from the list.
   - Click "Run."

Remember that macros can have a significant impact on your Outlook data, so be cautious when running them. Only run macros from trusted sources.

That's how you enable and create a macro in Microsoft Outlook using VBA. You can create more complex macros to automate various tasks within Outlook, such as sending automated emails, processing incoming messages, or managing your mailbox.



# Automatic scan of incoming mail
To automatically invoke a script in Outlook when an email arrives, you can use the built-in "Run a Script" rule feature. Here are the steps to set up a rule to run your script when an email arrives:

1. **Open Outlook:** Ensure that Outlook is open and running.

2. **Open Rules and Alerts:**
   - In Outlook 2010, 2013, 2016, or 2019: Click on the "File" tab, then click "Manage Rules & Alerts."
   - In Outlook 365 or Outlook 2021: Click on the "File" tab, then click "Manage Rules & Alerts."

3. **Create a New Rule:**
   - Click on "New Rule" to create a new rule.

4. **Start from a Blank Rule:**
   - In the "Rules Wizard" dialog, choose "Apply rule on messages I receive" and click "Next."

5. **Conditions (Optional):** If you want to apply the rule based on specific conditions (e.g., specific sender or subject), select the conditions. Otherwise, leave it blank and click "Next."

6. **Choose an Action:**
   - In the "What do you want to do with the message?" section, check "run a script."
   - Click on the "a script" link in the bottom panel.

7. **Select a Script:**
   - In the "Run Script" dialog, choose your script (the script you want to run when an email arrives).
   - If your script is not listed, click "Browse" to locate and select it.
   - Click "Next."

8. **Apply Rule to Messages:**
   - Review the rule description and click "Next."
   - Optionally, you can add any exceptions to the rule, or you can leave it as "no exceptions."
   - Click "Next."

9. **Name the Rule:**
   - Give your rule a name (e.g., "Run My Script on Incoming Emails").
   - Choose whether to run the rule on existing messages in the inbox (recommended if you want to apply the rule retroactively).
   - Click "Finish."

10. **Apply the Rule:**
    - In the "Rules and Alerts" dialog, review the rule you've created.
    - If everything looks correct, click "Apply" and then "OK."

Your rule is now set up to run your script when an email meeting the specified conditions arrives in your inbox. If you want to apply the rule to existing messages in your inbox, select the option to do so when creating the rule.
