#!/bin/sh
# 分割した zip を結合して xlCode-src-win.zip を作る
cd "$(dirname "$0")" && cat xlCode-src-win.zip.part* > xlCode-src-win.zip && echo "xlCode-src-win.zip を作成しました"
