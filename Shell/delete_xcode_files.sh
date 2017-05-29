<<COMMENT
Xcodeで溜まっていくファイルを削除するためのシェルスクリプト。
削除するフォルダは以下を参考にした。
http://dev.classmethod.jp/smartphone/remove-xcode8-related-unnecessary-files/
COMMENT

# Delete Archives
ls -l ~/Library/Developer/Xcode/Archives/
rm -rf ~/Library/Developer/Xcode/Archives/*

# Delete DerivedData
ls -l ~/Library/Developer/Xcode/DerivedData/
rm -rf ~/Library/Developer/Xcode/DerivedData/*

# Delete Device Support
ls -l ~/Library/Developer/Xcode/iOS\ DeviceSupport/
rm -rf ~/Library/Developer/Xcode/iOS\ DeviceSupport/*

# Delete Device Logs
ls -l ~/Library/Developer/Xcode/iOS\ Device\ Logs/
rm -rf ~/Library/Developer/Xcode/iOS\ Device\ Logs/*
