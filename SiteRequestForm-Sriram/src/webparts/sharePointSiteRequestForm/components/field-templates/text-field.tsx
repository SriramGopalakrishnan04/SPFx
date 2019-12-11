import * as React from 'react';
import TextField from '@material-ui/core/TextField';

const textFieldCss = {
    // width: '80%',
    minWidth: '200px',
    marginLeft: '8px',
    marginRight: '8px'
};

// This is necessary due to some other styles that are loaded into SharePoint pages causing the backgrounds of the fields to turn white.
const textFieldInputStyles = {
    'color': 'inherit',
    'backgroundColor': 'rgba(0,0,0,0)',
    'borderColor': 'inherit'
} as React.CSSProperties;

interface TextFieldTemplateProps {
    label: string;
    placeHolder: string;
    required?: boolean;
    multiline?: boolean;
    onChangeHandler: (fieldName: string, fieldValue: string) => void;
}

type DefaultProps = {
    required: boolean,
    multiline: boolean
};


const initialState = {
    inputValue: ''
};

type State = Readonly<typeof initialState>;

class TextFieldTemplate extends React.Component<TextFieldTemplateProps> {
    public readonly state: State = initialState;

    public static defaultProps: DefaultProps = {
        required: false,
        multiline: false
    };

    public render() {
        let errorState = false;

        if (this.props.required && this.state.inputValue.length === 0) {
            errorState = true;
        }

        return (
            <TextField
                // This is whack but you have to specify inputProps inside of inputProps to affect the actual 'input'
                InputProps={{inputProps: {style: textFieldInputStyles}}}
                error={errorState}
                required={this.props.required}
                multiline={this.props.multiline}
                fullWidth
                color="primary"
                onChange={(evt) => {
                        this.props.onChangeHandler(this.props.label, evt.target.value);
                        this.setState({inputValue: evt.target.value});
                    }
                }
                label={this.props.label}
                style={textFieldCss}
                placeholder={this.props.placeHolder}
                margin="normal"
                variant="outlined"
            />
        );
    }
}

export default TextFieldTemplate;