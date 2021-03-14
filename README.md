# Outlook MSG Viewer
There is is apparently a need to be able to open Microsoft Outlook Message files (*.msg) in Microsoft Outlook on 
Mac OS systems.

There is an easy trick.

* Install brew if not already installed.
* Run brew install python (which now installs python 3.9).
* Checkout this repository.
* Run pip3 install -r requirements.txt --user.
* Using Automator create an application which can open an MSG File in Outlook (Run Shell Script), call it MsgViewer.
* Using Automator create an application which can run multiple instances of the bofore created Application (Run Apple Script), call it OutlookPreview.


### How to create these applications:
### MsgViewer
```console
#!/bin/bash
runpy3 () {
/usr/local/bin/python3 <<'EOF' - "$@"
Place here the content of outlook_preview.py
EOF
}
runpy3 "$@"
```
### OutlookPreview
```console
on run {input, parameters}
    set f to quoted form of (POSIX path of (input as text))
    do shell script "source ~/.bash_profile"
    do shell script "open -n -a MsgViewer.app --args " & f
end run
```
Copy both applications into the Applications folder.
Associate the "msg" file extension with the OutlookPreview.



