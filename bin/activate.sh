# add the dev directory to the path and pythonpath
if [ $_OFF2TXT_ACTIVATED ]; then
    echo 'Already activated'
else
    pwd=`pwd`
    export _PATH_BEFORE_OFF2TXT_ACTIVATE=$PATH
    export PATH=$pwd/dev:$PATH

    if [ ! -z $PYTHONPATH ]; then
        export _PYTHONPATH_BEFORE_OFF2TXT_ACTIVATE=$PYTHONPATH
        export PYTHONPATH=$pwd/dev:$PYTHONPATH
    else
        export _PYTHONPATH_BEFORE_OFF2TXT_ACTIVATE=0
        export PYTHONPATH=$pwd/dev
    fi

    export _OFF2TXT_ACTIVATED=1
fi
