/*
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    https://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/

// add menu to Sheet
function onOpen()
{
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('GCP')
      .addItem('Extract data from email', 'extractData')
      .addItem('Remove expired data', 'deleteRows')
      .addItem('Calculate TTL', 'calculateTTL')
      .addToUi();
}

function calculateTTL()
{
    try
    {
        var sheet = SpreadsheetApp.getActiveSheet();
        var cell = sheet.getRange("D4");
        cell.setFormula('=ARRAYFORMULA(IF(ISBLANK(C4:INDEX(C4:C,COUNTA(C4:C))), "", DATEDIF(TODAY(), C4:INDEX(C4:C,COUNTA(C4:C)), "D")))');
    }
    catch (e)
    {
        Logger.log(e.toString());
    }
}

function extractData()
{
    try
    {
        var sheet = SpreadsheetApp.getActiveSheet();
        var label = sheet.getRange(1, 2).getValue();

        var labels = GmailApp.getUserLabelByName(label);
        var threads = labels.getThreads();

        var data = [];

        for (var i = 0; i < threads.length; i++)
        {
            // get first message from the thread
            var message = threads[i].getMessages()[0],
                content = message.getPlainBody(),
                tmp;

            // implement parsing rules using regular expressions
            if (content)
            {
                tmp = content.match(/project\s*(.*)was/);
                var project = tmp[1];

                tmp = content.match(/on\s*((\d{4})-(\d{2})-(\d{2})T(\d{2})\:(\d{2})\:(\d{2})[+-](\d{2})\:(\d{2}))/);
                var delOn = new Date(tmp[1]);

                tmp = content.match(/by\s*((\d{4})-(\d{2})-(\d{2})T(\d{2})\:(\d{2})\:(\d{2})[+-](\d{2})\:(\d{2}))/);
                var shutOn = new Date(tmp[1]);

                data.push([project, delOn, shutOn]);

                // delete message
                message.moveToTrash();
            }
        }

        // write to spreadsheet
        sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
    }
    catch (e)
    {
        Logger.log(e.toString());
    }
}

function deleteRows()
{
    try
    {
        var sheet = SpreadsheetApp.getActiveSheet();
        var range = sheet.getRange("D:D");
        var values = range.getValues();

        for (var i = values.length; i > 0; i--)
        {
            if (values[0, i] == '#NUM!' || (values[0, i] == 0 && values[0, i] != ""))
            {
                sheet.deleteRow(i + 1);
            }
        }
    }
    catch (e)
    {
        Logger.log(e.toString());
    }
}