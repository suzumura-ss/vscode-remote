Handle vscode-remote://
===

vscode-remote:// スキーマを処理するためのスクリプトです。
WSL2環境下の VSCode からの vscode-remote://wsl+(distro)/path/to/content を受け取ると、それをエクスプローラーで開きます。

WSL2環境下の VSCode + [Marp for VS Code](https://marketplace.visualstudio.com/items?itemName=marp-team.marp-vscode) でHTML/PDFプレビューを開けるようにするために作成しました。

install
---

    cscript vscode-remote.vbs --install

uninstall
---

    cscript vscode-remote.vbs --uninstall
