if [ $_OFF2TXT_ACTIVATED ]; then
    
    if [ $_PATH_BEFORE_OFF2TXT_ACTIVATE ]; then
        export PATH=$_PATH_BEFORE_OFF2TXT_ACTIVATE
    fi

    if [ $_PYTHONPATH_BEFORE_OFF2TXT_ACTIVATE ]; then
        if [  $_PYTHONPATH_BEFORE_OFF2TXT_ACTIVATE != 0 ]; then
            export PYTHONPATH=$_PYTHONPATH_BEFORE_OFF2TXT_ACTIVATE
        else
            unset PYTHONPATH
        fi
    fi

    unset _PATH_BEFORE_OFF2TXT_ACTIVATE
    unset _PYTHONPATH_BEFORE_OFF2TXT_ACTIVATE
    unset _OFF2TXT_ACTIVATED
else
    echo 'Not active'
fi
    