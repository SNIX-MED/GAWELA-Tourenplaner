Place the extracted Microsoft Edge WebView2 Fixed Version Runtime in this folder
before building the packaged app.

Expected result:
- assets\webview2\msedgewebview2.exe
- assets\webview2\...\additional runtime files

Alternative:
- set the environment variable GAWELA_WEBVIEW2_RUNTIME_DIR to the extracted
  runtime directory before running PyInstaller
