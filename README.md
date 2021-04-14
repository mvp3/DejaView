# Deja View

Deja View is a nifty lightweight Microsoft Word Add-In (VSTO) that automatically retains a document's 
view settings internally.

## Introduction

Deja View uses the [Custom XML Parts](https://docs.microsoft.com/en-us/visualstudio/vsto/custom-xml-parts-overview?view=vs-2019) 
embedding feature of Microsoft Word to save the application's view parameters within the document.
When a document is opened Deja View looks for any previously embedded view parameters. If found, Deja View attempts to restore the view parameters of the 
Microsoft Word application window.

#### User Interface
The Deja View add-in has one user interface (UI) component, namely a ribbon group. If the add-in 
is installed and loads correctly at the startup of Microsoft Word, it will be accessible under
the "Add-ins" tab of the main ribbon.

![Ribbon (Dark)](https://dejaview.lexem.cc/images/ribbon_dark.png)
 
##### Enable
This option allows a quick and easy means to temporarily enable / disable the Deja View add-in. 

##### Location
This option determines if Deja View will remember the document window's location. Default is checked. 
This allows users to reopen several documents, each returning to the same location on their screen.

This feature also allows for multiscreen displays. If a document is opened on a computer with a different 
number of display screens or a smaller resolution then the computer that the document was last saved on, 
so that the document would not be visible when displayed, Deja View will automatically detect that the 
document window is not visible on the present display and center the document on the primary screen.

##### Clear View Tags
If this button is pressed, all tags (custom XML parts) from Deja View will be remove from the active document.
The button will be displayed as disabed when no Deja View tags are discovered. If the document is saved
with Deja View enabled, view tags will be embedded and this button will immediately be enabled.

Mousing over this button pops up a super tooltip that displays all view parameters.

#### Custom XML Parts
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
    </application>
</lexidata>
```

## Installation

The **Dejaview.dll**, **Dejaview.dll.manifest**, and **Dejaview.vsto** files have all been packaged by an 
installation project built using [Wix toolset for Visual Studio](https://wixtoolset.org/).

#### Binary
The latest MSI installation binary may be downloaded here: [https://dejaview.lexem.cc/latest](https://dejaview.lexem.cc/latest)

#### Built With
Visual Studio Community 2019

#### Tested On
Microsoft Word for Microsoft Office 365 MSO (16.0.13628.20234) 64-bit

## Certificate Information
Deja View is signed with a standard SHA256RSA certificate. 
But the certificate authority (CA) root certificate used is not trusted by default
because it is not in the Trusted Root Certification Authorities store.

This is entirely a financial matter. If there is support for the project
I would be happy to purchase a trusted signing certificate from a 
public CA.

The present CA Root certificate is issued by [TBC CA](https://trinitybiblechurch.org/TBC-root-authority.crt)

## Support
Extremely limited support will be offered at this time.

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.

## License
Copyright 2021 M. V. Pereira

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

Manny Pereira - truthbearer@gmail.com

Project Link: https://github.com/mvp3/DejaView

## Acknowledgements
The original author of this project and its code is M. V. Pereira.

The code is very simple and much help was derived from online 
forums and [Visual Studio Docs](https://docs.microsoft.com/en-us/visualstudio/?view=vs-2019).
