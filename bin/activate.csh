# add the dev directory to the path and pythonpath
if ( ${?_OFF2TXT_ACTIVATED} ) then
    echo 'Already activated'
else
    set pwd=`pwd`
    setenv _PATH_BEFORE_OFF2TXT_ACTIVATE $PATH
    setenv PATH $pwd/dev:$PATH

    if ( ${?PYTHONPATH} ) then
        setenv _PYTHONPATH_BEFORE_OFF2TXT_ACTIVATE $PYTHONPATH
        setenv PYTHONPATH $pwd/dev:$PYTHONPATH
    else
        setenv _PYTHONPATH_BEFORE_OFF2TXT_ACTIVATE
        setenv PYTHONPATH $pwd/dev
    endif

    setenv _OFF2TXT_ACTIVATED
endif
