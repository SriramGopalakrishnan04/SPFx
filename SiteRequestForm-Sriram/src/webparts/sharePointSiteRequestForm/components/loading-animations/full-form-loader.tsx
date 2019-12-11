import * as React from 'react';
import CircularProgress from '@material-ui/core/CircularProgress';
import Check from '@material-ui/icons/Check';
import Warning from '@material-ui/icons/Warning';
import Fade from '@material-ui/core/Fade';
import Zoom from '@material-ui/core/Zoom';
import green from '@material-ui/core/colors/green';
import red from '@material-ui/core/colors/red';

const fullFormContainerHiddenCss = {
    width: '100%',
    height: '100%',
    position: 'absolute',
    zIndex: 99,
    display: 'flex',
    opacity: 0,
    transition: 'opacity 0.3s ease 0s',
    flexDirection: 'column',
    pointerEvents: 'none'
} as React.CSSProperties;

const fullFormContainerVisibleCss = {
    width: '100%',
    height: '100%',
    position: 'absolute',
    zIndex: 99,
    display: 'flex',
    background: 'rgba(51,51,51,0.5)',
    opacity: 1,
    transition: 'opacity 0.3s ease 0s',
    flexDirection: 'column'
} as React.CSSProperties;

const fullFormContainerCompleteCss = {
    width: '100%',
    height: '100%',
    position: 'absolute',
    zIndex: 99,
    display: 'flex',
    background: green[500],
    opacity: 1,
    transition: 'opacity 0.3s ease 0s, background 0.3s ease 0s',
    flexDirection: 'column'
} as React.CSSProperties;

const fullFormContainerWarningCss = {
    width: '100%',
    height: '100%',
    position: 'absolute',
    zIndex: 99,
    display: 'flex',
    background: red[500],
    opacity: 1,
    transition: 'opacity 0.3s ease 0s, background 0.3s ease 0s',
    flexDirection: 'column'
} as React.CSSProperties;

const spinnerCss = {
    margin: 'auto',
    zIndex: 101
} as React.CSSProperties;

const checkCss = {
    margin: 'auto',
    width: '50%',
    height: '50%',
    flex: 3,
    zIndex: 101
} as React.CSSProperties;

const messageCss = {
    flex: 1,
    textAlign: 'center',
    zIndex: 101
} as React.CSSProperties;

const getCssProps = (props) => {
    if (props.warning) {
        return fullFormContainerWarningCss;
    } else if (props.complete) {
        return fullFormContainerCompleteCss;
    } else if (props.active) {
        return fullFormContainerVisibleCss;
    } else {
        return fullFormContainerHiddenCss;
    }
};

const FullFormLoader = (props) => {
    const currentCssProps = getCssProps(props);
    return (
        <div>
            <Fade in={!props.complete && props.active} mountOnEnter={true} unmountOnExit={true}>
                <div style={currentCssProps}>
                    <Fade in={!props.complete && props.active}>
                        <CircularProgress style={spinnerCss} size={100} />
                    </Fade>
                </div>
            </Fade>
            <Fade in={props.complete} mountOnEnter={true} unmountOnExit={true}>
                <div style={currentCssProps}>
                    <Zoom in={props.complete}>
                        <Check color="primary" style={checkCss} />
                    </Zoom>
                </div>
            </Fade>
            <Fade in={props.warning} mountOnEnter={true} unmountOnExit={true}>
                <div style={currentCssProps}>
                    <Zoom in={props.warning}>
                        <Warning color="primary" style={checkCss} />
                    </Zoom>
                    <Fade in={props.warning}>
                        <div style={messageCss}>
                            {props.warningMessage}
                        </div>
                    </Fade>
                </div>
            </Fade>
        </div>
    );
};

export default FullFormLoader;