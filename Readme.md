# PowerShell Script: Extract Email Headers from Outlook Message Files

## Overview

This PowerShell script is designed to extract email headers from Outlook message files (.msg) within a specified directory and its subdirectories. It utilizes the Microsoft Outlook COM object to accomplish this task.

## Prerequisites

Before running this script, ensure that you have the following prerequisites in place:

- **Windows OS**: This script is intended to run on Windows due to its dependency on the Outlook COM object.
- **Microsoft Outlook**: Microsoft Outlook must be installed on the machine where the script is executed.
- **PowerShell**: You should have PowerShell installed. This script is compatible with PowerShell 5.1 and above.

## Usage

1. Clone or download this repository to your local machine.

2. Open a PowerShell terminal.

3. Navigate to the directory where you have saved the script.

4. Run the script using the following command:

```powershell
.\ExtractHeaders.ps1
```

5. The script will start processing and extract email headers from all .msg files found in the specified directory and its subdirectories.

6. The extracted email headers will be saved to a file named HeaderOutput.txt in the same directory as the script.

## Sample Output

Here is an example of the extracted email headers in HearderOutput.txt:

```
Headers from: C:\path\to\example.msg

Received: from mail.example.com (mail.example.com [203.0.113.1])
    by mail.example.com (Postfix) with ESMTP id ABCDEFGHIJKL
    for <recipient@example.com>; Fri, 01 Jan 2023 12:34:56 +0000 (UTC)

```

## License

This script is licensed under the MIT License. See the LICENSE file for details.

## Author

[calebchadz](https://github.com/calebchadz)

## Acknowledgments

Thanks to Microsoft for providing the Outlook COM object for PowerShell.


