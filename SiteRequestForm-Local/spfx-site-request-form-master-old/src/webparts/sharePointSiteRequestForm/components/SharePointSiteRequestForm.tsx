import * as React from 'react';

import { MuiThemeProvider, createMuiTheme } from '@material-ui/core/styles';
import Paper from '@material-ui/core/Paper';

import amber from '@material-ui/core/colors/amber';
import grey from '@material-ui/core/colors/grey';
// import blue from '@material-ui/core/colors/blue';
// import yellow from '@material-ui/core/colors/yellow';
import green from '@material-ui/core/colors/green';

import Button from '@material-ui/core/Button';

import TextFieldTemplate from './field-templates/text-field';

import PeoplePickerTemplate from './field-templates/people-picker-template';

import FullFormLoader from './loading-animations/full-form-loader';

import styles from './SharePointSiteRequestForm.module.scss';
import { ISharePointSiteRequestFormProps } from './ISharePointSiteRequestFormProps';

import { createSiteRequest, getListItemEntityTypeName } from '../services/sp-rest';

const blueTheme = createMuiTheme({
  palette: {
    primary: amber,
    secondary: grey,
    type: "dark"
  },
});

const getItemDataForPost = (formData: any) => {
  let itemData = {
    Title: formData["Team Name"],
    PrimaryOwner: formData["Primary Owner"],
    SecondaryOwner: formData["Secondary Owner"],
    AdditionalOwners: formData["Additional Owners"],
    Members: formData["Members"]
  };

  return itemData;
};

const containerCss = {
  display: 'flex',
  flexWrap: 'wrap'
} as React.CSSProperties;

const buttonWrapperCss = {
  margin: 8,
  position: 'relative',
} as React.CSSProperties;

const buttonProgressCss = {
  color: green[500],
  position: 'absolute',
  top: '50%',
  left: '50%',
  marginTop: -12,
  marginLeft: -12,
} as React.CSSProperties;

const buttonSuccessCss = {
  backgroundColor: green[500],
  margin: 8
};

const buttonBaseCss = {
  margin: 8
};

const successMessageCss = {
  color: green[500]
};

const initialState = {
  isSubmitted: false,
  isValidForm: false,
  errorMessage: "Please provide a value for all required fields.",
  formState: "missingData",
  formData: {
    "Team Name": "",
    "Primary Owner": [],
    "Secondary Owner": [],
    "Additional Owners": [],
    "Members": []
  },
  loading: false,
  isConfigured: false
};

type State = Readonly<typeof initialState>;

export default class SharePointSiteRequestForm extends React.Component<ISharePointSiteRequestFormProps, {}> {
  public readonly state: State = initialState;

  private getButtonCss = () => {
    if (this.state.isSubmitted) {
      return buttonSuccessCss;
    } else {
      return buttonBaseCss;
    }
  }

  public handleTextChange = (fieldName: string, fieldValue: string) => {
    this.setState({
      formData: { ...this.state.formData, [fieldName]: fieldValue }
    });
  }

  public handleUserFieldChange = (fieldName: string, fieldValue: any[]) => {
    this.setState({
      formData: { ...this.state.formData, [fieldName]: fieldValue }
    });
  }

  // // Handle some setup after the component mounts
  public componentDidMount() {
    const listName = this.props.listName;

    if (listName && listName.trim()) {
      this.setState({
        isConfigured: true
      });
    }
  }

  public componentDidUpdate(prevProps, prevState, snapshot) {
    const listName = this.props.listName;

    if (listName && listName.trim() && !this.state.isConfigured) {
      this.setState({
        isConfigured: true
      });
    } else if (this.state.isConfigured && !(listName && listName.trim())) {
      this.setState({
        isConfigured: false
      });
    }
  }

  private setOwnerConflictErrorMessage() {
    this.setState({
      errorMessage: "Primary Owner cannot be the same as Secondary Owner",
      formState: "ownersMatch"
    });
  }

  private setMissingDataMessage() {
    this.setState({
      errorMessage: "Please provide a value for all required fields.",
      formState: "missingData"
    });
  }

  private clearErrorState() {
    this.setState({
      errorMessage: "",
      formState: "clean"
    });
  }

  private validateFormData() {
    const isMissingRequiredData = this.state.formData["Team Name"].length === 0 || this.state.formData["Primary Owner"].length !== 1 || this.state.formData["Secondary Owner"].length !== 1;
    const hasMatchingPrimarySecondary = (this.state.formData["Primary Owner"].length && this.state.formData["Secondary Owner"].length && (this.state.formData["Primary Owner"][0]["Key"] === this.state.formData["Secondary Owner"][0]["Key"]));
    if (isMissingRequiredData) {
      if (this.state.formState !== "missingData") {
        this.setMissingDataMessage();
      }
    } else if (hasMatchingPrimarySecondary) {
      if (this.state.formState !== "ownersMatch") {
        this.setOwnerConflictErrorMessage();
      }
    } else if (this.state.formState !== "clean") {
      this.clearErrorState();
    }
  }

  private handleRequestButtonClick() {
    // Set the state as loading before trying to create the request.

    if (this.state.formState === "clean") {
      this.setState({
        loading: true
      });

      let itemData = getItemDataForPost(this.state.formData);
      // Try to create the request.
      createSiteRequest(this.props.webpartContext, itemData, this.props.listName).then((result) => {
        this.setState({
          loading: false,
          isSubmitted: true
        });
      }).catch(err => {
        this.setState({
          loading: false,
          didError: true
        });
      });
    }
  }

  private renderFormBody() {
    // if (this.state.isSubmitted) {
    //   return (

    //   );
    // } else {
    return (

      <form style={containerCss} noValidate autoComplete="off">
        {this.state.isSubmitted && <div style={successMessageCss}>Your site request has been submitted.</div>}
        {!this.state.isSubmitted && <br />}
        <TextFieldTemplate label="Team Name" placeHolder="E.G. IS Web Content Management" onChangeHandler={this.handleTextChange} required />

        <PeoplePickerTemplate label={"Primary Owner"} wpContext={this.props.webpartContext} onChangeHandler={this.handleUserFieldChange} required singleValue error={this.state.formState === "ownersMatch"}/>
        <PeoplePickerTemplate label={"Secondary Owner"} wpContext={this.props.webpartContext} onChangeHandler={this.handleUserFieldChange} required singleValue error={this.state.formState === "ownersMatch"}/>

        <PeoplePickerTemplate label={"Additional Owners"} wpContext={this.props.webpartContext} onChangeHandler={this.handleUserFieldChange} />

        <PeoplePickerTemplate label={"Members"} wpContext={this.props.webpartContext} onChangeHandler={this.handleUserFieldChange} />

        <br />
        <div style={buttonWrapperCss}>
          <Button style={this.getButtonCss()} disabled={!(this.state.formState === "clean") || this.state.loading || this.state.isSubmitted} variant="outlined" color="primary" onClick={() => { this.handleRequestButtonClick(); }}>Submit</Button>
          {/* {this.state.loading && <CircularProgress size={24} style={buttonProgressCss} />} */}
        </div>
        {/* {this.state.loading && <FullFormLoader size={24} style={buttonProgressCss} />} */}

      </form>
    );
    // }
  }

  public render(): React.ReactElement<ISharePointSiteRequestFormProps> {
    this.validateFormData();
    // if (this.state.isLoadingTypeName) {
    return (
      <MuiThemeProvider theme={blueTheme}>

        {/* <div> */}
        <div className={styles.oneDriveForm}>
          <Paper>
            <FullFormLoader active={this.state.loading} complete={this.state.isSubmitted} warning={!this.state.isConfigured} warningMessage="The webpart is not configured" />

            {/* <div> */}
              {/* <div> */}
              <div className={styles.row}>
                {/* <div> */}
                <div className={styles.column}>
                  <span className={styles.title}>SharePoint Team Site Request</span>

                  <p className={styles.description}>{this.state.errorMessage}&nbsp;</p>


                  {this.renderFormBody()}

                </div>
              </div>
          </Paper>

        </div>
      </MuiThemeProvider>
    );

  }
}
