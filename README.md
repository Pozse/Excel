<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel VBA Automation</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 0;
            background-color: #f4f4f4;
            color: #333;
        }
        .container {
            width: 80%;
            margin: 50px auto;
            background: #fff;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0px 0px 10px rgba(0, 0, 0, 0.1);
        }
        h1, h2 {
            color: #0056b3;
        }
        h1 {
            text-align: center;
        }
        ul {
            list-style-type: square;
            padding-left: 20px;
        }
        code {
            background: #eee;
            padding: 2px 5px;
            border-radius: 3px;
            font-family: monospace;
        }
        .note {
            background: #fff3cd;
            padding: 10px;
            border-left: 5px solid #ffa502;
            margin: 15px 0;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Excel VBA Automation – Data Sync & Conditional Formatting</h1>
        
        <h2>Overview</h2>
        <p>This VBA script automates data synchronization and conditional formatting between two sheets: <strong>"EFC Master"</strong> and <strong>"Device Information"</strong>. It ensures efficient data transfer and enforces structured user interactions.</p>
        
        <h2>Features</h2>
        <ul>
            <li><strong>Auto Data Sync:</strong> Selecting any cell in <code>B2:Q14</code> on <strong>"Device Information"</strong> copies values and formatting to <strong>"EFC Master"</strong>.</li>
            <li><strong>Conditional Formatting:</strong>
                <ul>
                    <li>If <code>Q9</code> changes to <code>"No"</code>, it turns <strong style="color:red;">red</strong> and prompts the user about the <strong>Go-Back</strong> sheet.</li>
                    <li>If <code>"No"</code> is selected, a hidden hyperlink (<code>Q273</code>) automatically opens the <strong>Go-Back</strong> sheet.</li>
                    <li>If <code>"Yes"</code>, <code>Q9</code> turns <strong style="color:green;">green</strong>.</li>
                    <li><code>Q9</code> is always set to <strong>"Non-Managed CC Switch Resolved."</strong></li>
                </ul>
            </li>
            <li><strong>Optimized Performance:</strong> Uses <em>screen updating and event disabling</em> to prevent unnecessary recalculations.</li>
        </ul>
        
        <h2>How to Use</h2>
        <ol>
            <li><strong>Enable Macros:</strong> Ensure macros are enabled in Excel (<code>File > Options > Trust Center > Macro Settings</code>).</li>
            <li><strong>Selection-Based Sync:</strong> Click on any cell in <code>B2:Q14</code> on <strong>"Device Information"</strong> to transfer data to <strong>"EFC Master"</strong>.</li>
            <li><strong>Monitor Q9:</strong> If <code>Q9</code> changes, follow prompts for the <strong>Go-Back</strong> sheet when needed.</li>
        </ol>
        
        <h2>Customization</h2>
        <ul>
            <li>Modify <code>B2:Q14</code> if different data ranges need syncing.</li>
            <li>Change <code>Q9</code> if a different cell should control formatting.</li>
            <li>Adjust <code>Q273</code> if the Go-Back hyperlink is stored elsewhere.</li>
        </ul>
        
        <div class="note">
            <strong>Note:</strong>
            <ul>
                <li>Ensure the <strong>Go-Back sheet hyperlink</strong> exists in <code>Q273</code> or update the reference.</li>
                <li>Undo (<code>CTRL + Z</code>) won’t work, as VBA directly modifies cells.</li>
                <li>If macros don’t run, check <strong>macro security settings</strong> and enable VBA.</li>
            </ul>
        </div>
        
        <h2>License</h2>
        <p>This script is provided as-is. Modify and use it as needed.</p>
    </div>
</body>
</html>
