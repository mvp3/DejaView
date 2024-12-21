# Deja View

Deja View is a nifty lightweight Microsoft Word Add-In (VSTO) that automatically retains a document's 
view settings internally.

## Introduction

Deja View uses the [Custom XML Parts](https://docs.microsoft.com/en-us/visualstudio/vsto/custom-xml-parts-overview?view=vs-2019) 
embedding feature of Microsoft Word to save the application's view parameters within the document.
When a document is opened Deja View looks for any previously embedded view parameters. If found, Deja View attempts to restore the view parameters of the 
Microsoft Word application window.

### User Interface
The Deja View add-in has one main user interface (UI) component, namely a ribbon group.
An option dialog is made available from the Deja View ribbon group.

#### Ribbon Group
If the add-in is installed and loads correctly at the startup of Microsoft Word, 
it will be accessible under the "Add-ins" tab of the main ribbon.

![Ribbon (Dark)](https://dejaview.lexem.cc/images/ribbon_dark.png)
 
##### Clear View Tags
If this button is pressed, all tags (custom XML parts) from Deja View will be remove from the active document.
The button will be displayed as disabed when no Deja View tags are discovered. If the document is saved
with Deja View enabled, view tags will be embedded and this button will immediately be enabled.

##### Check for Update
Use this button to manually check for updates to the Deja View add-in.

Also, mousing over this button will show a tooltip that displays the current version of Deja View.

##### Options
This button will show the Deja View options dialog.

#### Options Dialog

![Options Dialog](https://dejaview.lexem.cc/images/options_dialog_1.0.3.png)

##### Enable Deja View
This option allows a quick and easy means to temporarily enable / disable the Deja View add-in. 

##### Prompt before saving view settings
If checked, Deja View will ask before saving view settings to this document. 

##### Automatically check for updates
If checked, Deja View will automatically check for updates when the add-in is loaded. It will not check more than once per day.

##### View Tags
This button shows a dialog that displays the Deja View tags currently embedded in the active document.

##### Current
This button shows a dialog that displays the current view parameters. If Deja View is enabled, 
these parameters will be saved to the document when the document is saved.

##### Apply Last View
This button applies the last Deja View view parameters to the current document. This should automatically happen when a document 
is loaded. This button simply offers a manual method.

##### Defaults
This button resets all options to the default. It prompts before performing the reset.

##### View Logs
This button shows a dialog that displays the Deja View logs that correspond to the active document.

##### Update URL
The Internet URL for Deja View to use when checking for updates. 
The default is https://dejaview.lexem.cc/autoupdate.

##### Window Location
This option determines if Deja View will remember the document window's location. Default is checked. 
This allows users to reopen several documents, each returning to the same location on their screen.

This feature also allows for multiscreen displays. If a document is opened on a computer with a different 
number of display screens or a smaller resolution then the computer that the document was last saved on, 
so that the document would not be visible when displayed, Deja View will automatically detect that the 
document window is not visible on the present display and center the document on the primary screen.

##### Navigation Pane
Save and restore the visibility and width details for side navigation pane.

##### View Layout
Save and restore the document window view layout. The value corresponds to Microsoft's [WdViewType Enum](https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdviewtype?view=word-pia).

##### Zoom
Save and restore the document zoom level.

##### Ruler
Save and restore the setting to show rulers.

##### Ribbon
Save and restore the settings for how the ribbon should be displayed. Typically the ribbon is expanded 
and pinned. But users may wish to hide the ribbon when viewing certain documents.

### Custom XML Parts
Invisibly and quietly, Deja View embeds a small XML segment (called "tags" in the user interface) 
into the XML-based Word document (.docx) when it is saved. Changes to application window view 
dimensions will only be retained if the document is saved.

The following is an example of an embedded XML segment:

```xml
<lexidata xmlns="Dejaview">
    <navigation>
        <width>231</width>
        <show>true</show>
    </navigation>
    <application>
        <left>3</left>
        <top>3</top>
        <width>700</width>
        <height>788</height>
        <windowstate>0</windowstate>
        <view>1</view>
        <draft>true</draft>
        <rulers>true</rulers>
        <zoom>100</zoom>
        <ribbonheight>161</ribbonheight>
        <location>
            <dauid>4297f44b13955235245b2497399d7a93</dauid>
            <uid></uid>
            <top></top>
            <left></left>
            <ts></ts>
        </location>
    </application>
</lexidata>
```

## Installation

The **Dejaview.dll**, **Dejaview.dll.manifest**, and **Dejaview.vsto** files have all been packaged by an 
installation project built using the Microsoft Installer tool.

#### Binary
The latest installation binary may be downloaded here: [https://dejaview.lexem.cc/latest](https://dejaview.lexem.cc/latest)

#### Built With
Visual Studio Community 2022

#### Tested On
Microsoft® Word for Microsoft 365 MSO (Version 2308 Build 16.0.16731.20052) 64-bit 

## Certificate Information
Deja View is signed with a standard SHA256RSA certificate issued by
Sectigo Public Code Signing CA R36. It is valid from 1/31/2023 to 2/1/2024.

If there is no support for the project I will be forced to resort 
to an untrusted self-signed certificate.

## Support
Extremely limited support will be offered at this time.

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.

## License
Copyright 2023 M. V. Pereira

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

## Contact

M. V. Pereira - truthbearer@gmail.com

Project Link: https://github.com/mvp3/DejaView

Public Site: https://dejaview.lexem.cc

## Acknowledgements
The original author of this project and its code is M. V. Pereira.

The code is very simple and much help was derived from online 
forums and [Visual Studio Docs](https://docs.microsoft.com/en-us/visualstudio/?view=vs-2022).
