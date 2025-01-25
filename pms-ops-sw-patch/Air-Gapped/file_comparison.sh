#!/bin/bash

# SFTP 서버 정보
SFTP_SERVER="SFTP_SERVER"
PORT="PORT"
USERNAME="USERNAME"

# SFTP 접속
if sftp -oPort=$PORT $USERNAME@$SFTP_SERVER &>/dev/null <<EOF; then
pwd
bye
EOF
    echo "SFTP 서버에 접속 설공"
else
    echo "SFTP 서버에 접속 실패"
fi
bash
