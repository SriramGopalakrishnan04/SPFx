import * as React from 'react';
import Downshift from 'downshift';
import TextField from '@material-ui/core/TextField';
import Paper from '@material-ui/core/Paper';
import MenuItem from '@material-ui/core/MenuItem';
import Chip from '@material-ui/core/Chip';

import PeopleSearchService from '../services/people-search-svc';
import TimeoutHandler from '../services/timeout-service';

const styles = {
    root: {
        flexGrow: 1,
        width: '100%',
        marginLeft: 8,
        marginRight: 8
    } as React.CSSProperties,
    container: {
        flexGrow: 1,
        position: 'relative',
    } as React.CSSProperties,
    paper: {
        position: 'absolute',
        zIndex: 2,
        marginTop: 4,
        left: 0,
        right: 0,
        background: '#fff'
    } as React.CSSProperties,
    chip: {
        margin: `${12 / 2}px ${8 / 4}px`,
    } as React.CSSProperties,
    inputRoot: {
        flexWrap: 'wrap',
        flex: 1,
        minWidth: 150,
        color: 'inherit',
        'backgroundColor': 'rgba(0,0,0,0)',
        'borderColor': 'inherit'
    } as React.CSSProperties,
    inputInput: {
        width: 'auto',
        flexGrow: 1,
    } as React.CSSProperties,
    divider: {
        height: 8 * 2,
    } as React.CSSProperties,
    selectBox: {
        width: 'calc(100% - 14px)',
        flexWrap: 'wrap'
    } as React.CSSProperties
};

function renderInput(inputProps) {
    const { InputProps, classes, ref, ...other } = inputProps;

    return (
        <TextField
            // fullWidth
            margin="normal"
            variant='outlined'
            required={InputProps.required}
            InputProps={{
                inputRef: ref,
                style: classes.selectBox,
                ...InputProps,
            }}
            inputProps={{ style: classes.inputRoot }}
            {...other}
        />
    );
}

function renderSuggestion({ suggestion, index, itemProps, highlightedIndex, selectedItem }) {
    const isHighlighted = highlightedIndex === index;

    let isSelected = false;
    for (let item of selectedItem) {
        if (item.Key === suggestion.Key) {
            isSelected = true;
        }
    }

    if (!isSelected) {
        return (
            <MenuItem

                {...itemProps}
                key={suggestion.Key}
                selected={isHighlighted}
                component="div"
                style={{
                    fontWeight: isSelected ? 600 : 400,
                    color: 'rgba(0, 0, 0, 0.87)',
                    textOverflow: 'ellipsis',
                    display: 'block'
                }}
            >
                {`${suggestion.DisplayText} | ${suggestion.Description}`}
            </MenuItem>
        );
    } else {
        // If it is already selected, don't return any menu item

    }
}




const initialState = {
    inputValue: '',
    selectedItem: [],
    peopleSuggestions: [],
    disabled: false
};

type State = Readonly<typeof initialState>;

class DownshiftMultiple extends React.Component<{ required?: boolean; singleValue?: boolean; error?: boolean; peoplePickerService: PeopleSearchService; label: string; onChangeHandler: (fieldName: string, fieldValue: string[]) => void }> {
    public readonly state: State = initialState;

    // Get the suggestions from the peopleSuggestions state value. Limit to 5?
    private getSuggestions = (value) => {
        const inputValue = value.trim().toLowerCase();
        const inputLength = inputValue.length;
        let count = 0;

        return inputLength === 0
            ? []
            : this.state.peopleSuggestions.filter(suggestion => {
                // determine if we already have 5 suggestions or the current suggestion matches the search input.
                const keep = count < 5 && (suggestion.Description.slice(0, inputLength).toLowerCase() === inputValue || suggestion.DisplayText.slice(0, inputLength).toLowerCase() === inputValue || suggestion.Key.slice(0, inputLength).toLowerCase() === inputValue);

                if (keep) {
                    count += 1;
                }

                return keep;
            });
    }

    // Handle when the backspace key is pressed to remove the last 'selected' item when there is no text in the input.
    private handleKeyDown = event => {
        const { inputValue, selectedItem } = this.state;
        if (selectedItem.length && !inputValue.length && event.key === 'Backspace') {
            let item = selectedItem[selectedItem.length - 1];
            this.handleDelete(item)();
        }
    }

    private handleInputChange = event => {
        const inputVal = event.target.value;
        this.setState({ inputValue: event.target.value });

        if (inputVal.length >= 3) {
            TimeoutHandler.setTimeout('people-picker', () => {
                this.props.peoplePickerService.getSuggestions(inputVal).then(res => {
                    let pplResults = JSON.parse(res.value);
                    this.setState({
                        peopleSuggestions: pplResults
                    });
                });
            }, 1000);
        } else {
            TimeoutHandler.removeTimeout('people-picker');
        }
    }

    // Handle when there is a change to the 'selected' items
    private handleChange = item => {
        let { selectedItem } = this.state;

        if (this.props.singleValue) {
            selectedItem = [item];
        } else if (selectedItem.indexOf(item) === -1) {
            selectedItem = [...selectedItem, item];
        }

        this.props.onChangeHandler(this.props.label, selectedItem);

        this.setState({
            inputValue: '',
            selectedItem,
        });
    }

    private handleDelete = item => () => {
        const selectedItem = [...this.state.selectedItem];

        selectedItem.splice(selectedItem.indexOf(item), 1);

        this.props.onChangeHandler(this.props.label, selectedItem);

        this.setState({ selectedItem: selectedItem });
    }

    public render() {
        const classes = styles;

        const { inputValue, selectedItem } = this.state;

        let errorState = false;
        let inputDisabled = false;

        let placeholderVal = "Search for and select multiple users";
        if ((this.props.required && this.state.selectedItem.length === 0) || this.props.error) {
            errorState = true;
        }

        if (this.props.singleValue) {
            placeholderVal = "Search for and select one user";

            if (this.state.selectedItem.length !== 0 && !this.props.error) {
                inputDisabled = true;
            }
        }

        return (
            <Downshift
                id="downshift-multiple"
                inputValue={inputValue}
                // Handle when there is a change to the 'selected' items
                onChange={this.handleChange}
                selectedItem={selectedItem}
            >
                {({
                    getInputProps,
                    getItemProps,
                    isOpen,
                    inputValue: inputValue2,
                    selectedItem: selectedItem2,
                    highlightedIndex,
                }) => (
                        <div style={classes.container}>
                            {renderInput({
                                fullWidth: true,
                                classes,

                                InputProps: getInputProps({
                                    error: errorState,
                                    disabled: inputDisabled,
                                    required: this.props.required,
                                    startAdornment: selectedItem.map(item => {
                                        return (
                                            <Chip
                                                key={item.Description}
                                                tabIndex={-1}
                                                label={item.DisplayText}
                                                style={classes.chip}
                                                onDelete={this.handleDelete(item)}
                                            />
                                        );
                                    }),
                                    onChange: this.handleInputChange,
                                    onKeyDown: this.handleKeyDown,
                                    placeholder: placeholderVal
                                }),
                                label: this.props.label,
                            })}
                            {/* {true ? ( */}
                            {isOpen ? (
                                <Paper style={classes.paper} square>
                                    {this.getSuggestions(inputValue2).map((suggestion, index) =>
                                        renderSuggestion({
                                            suggestion,
                                            index,
                                            itemProps: getItemProps({ item: suggestion }),
                                            highlightedIndex,
                                            selectedItem: selectedItem2,
                                        }),
                                    )}
                                </Paper>
                            ) : null}
                        </div>
                    )}
            </Downshift>
        );
    }
}


function PeoplePickerTemplate(props) {
    const classes = styles;
    const peopleSearchSvc = new PeopleSearchService(props.wpContext);

    let singleValue = false;

    if (props.singleValue) {
        singleValue = true;
    }

    return (
        <div style={classes.root}>
            <DownshiftMultiple required={props.required} singleValue={singleValue} error={props.error} label={props.label} peoplePickerService={peopleSearchSvc} onChangeHandler={props.onChangeHandler} />
        </div>
    );
}

export default PeoplePickerTemplate;