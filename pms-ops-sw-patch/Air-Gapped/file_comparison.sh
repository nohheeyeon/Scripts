#!/bin/bash

# SSH 서버 정보
SSH_SERVER="SSH_SERVER"
PORT="PORT"
USERNAME="USERNAME"
REMOTE_DIRECTORY="DIRECTORY"

# 원격 서버의 패치셋에 있는 파일 이름 추출
ssh -p $PORT $USERNAME@SSH_SERVER "find $REMOTE_DIRECTORY -type f -or -type d" >remote_files.txt
if [ $? -ne 0 ]; then
    echo "원격 서버에서 파일/폴더 목록을 가져오는 데 실패"
    exit 1
fi
echo "원격 서버의 파일/폴더 목록 추출 완료"

# 추출된 파일/폴더 목록 출력
echo "원격 서버의 파일/폴더 목록 :"
cat remote_files.txt
bash
