if ( ${?_OFF2TXT_ACTIVATED} ) then
    
    if ( ${?_PATH_BEFORE_OFF2TXT_ACTIVATE} ) then
        setenv PATH $_PATH_BEFORE_OFF2TXT_ACTIVATE
    endif

    if ( ${?_PYTHONPATH_BEFORE_OFF2TXT_ACTIVATE} ) then
        if ( { eval 'test ! -z $_PYTHONPATH_BEFORE_OFF2TXT_ACTIVATE' } ) then
            setenv PYTHONPATH $_PYTHONPATH_BEFORE_OFF2TXT_ACTIVATE
        else
            unsetenv PYTHONPATH
        endif
    endif

    unsetenv _PATH_BEFORE_OFF2TXT_ACTIVATE
    unsetenv _PYTHONPATH_BEFORE_OFF2TXT_ACTIVATE
    unsetenv _OFF2TXT_ACTIVATED
else
    echo 'Not active'
endif
    