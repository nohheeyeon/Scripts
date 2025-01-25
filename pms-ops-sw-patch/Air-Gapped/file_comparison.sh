#!/bin/bash

# SSH 서버 정보
SSH_SERVER="SFTP_SERVER"
PORT="PORT"
USERNAME="USERNAME"
DIRECTORY="DIRECTORY"

# SSH를 통해 원격 명령 실행
ssh -p $PORT $USERNAME@$SSH_SERVER "find $DIRECTORY -type f -or -type d"
bash
