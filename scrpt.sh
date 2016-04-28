#!/bin/sh

# this file aut-corrects git commit blunders

git filter-branch --env-filter '
OLD_EMAIL="jccoki@development"
CORRECT_NAME="JC Campomanes"
CORRECT_EMAIL="john.campomanes@coredataresearch.com"
if [ "$GIT_COMMITTER_EMAIL" = "$OLD_EMAIL" ]
then
    export GIT_COMMITTER_NAME="$CORRECT_NAME"
    export GIT_COMMITTER_EMAIL="$CORRECT_EMAIL"
fi
if [ "$GIT_AUTHOR_EMAIL" = "$OLD_EMAIL" ]
then
    export GIT_AUTHOR_NAME="$CORRECT_NAME"
    export GIT_AUTHOR_EMAIL="$CORRECT_EMAIL"
fi
' --tag-name-filter cat -- --branches --tags
