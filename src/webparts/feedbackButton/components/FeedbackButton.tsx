import * as React from 'react';
import styles from './FeedbackButton.module.scss';
import { IFeedbackButtonProps } from './IFeedbackButtonProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IFeedbackButtonState } from './IFeedbackButtonState';
import { DefaultButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { IDigestCache, DigestCache } from '@microsoft/sp-http';
import axios from 'axios';

export default class FeedbackButton extends React.Component<IFeedbackButtonProps, IFeedbackButtonState> {

  constructor(props) {
    super(props);
    this.buttonClicked = this.buttonClicked.bind(this);
    this.onClosePanel = this.onClosePanel.bind(this);
    this.submitReport = this.submitReport.bind(this);
    this.state = {
      location: "",
      feedback: "",
      buttonVal: false,
      userEmail: "",
      thing: null
    }
  }

  public render(): React.ReactElement<IFeedbackButtonProps> {
    let disabled:any = this.state.buttonVal ? 'disabled' : null;
    return (
      <div className={ styles.feedbackButton }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <DefaultButton
              disabled={disabled}
              primary={true}
              text='Report Issue'
              onClick={ this.buttonClicked }
              />
              { this.state.buttonVal ? 
                 <Panel
                 isOpen={ this.state.buttonVal }
                 type={ PanelType.smallFixedFar }
                 onDismiss={ this.onClosePanel }
                //  onRenderFooterContent={ this._onRenderFooterContent }
                 headerText='Report an Issue with this page'
                 closeButtonAriaLabel='Close'>
                <TextField label="What would you like to report?" multiline rows={5} />
                <div className={styles.submitButton} style={{marginTop: 10 + 'px'}}>
                  <DefaultButton primary={true} text='Submit' onClick={this.submitReport}/>
                </div>
               </Panel>
              : null}
            </div>
          </div>
        </div>
      </div>
    );
  }

  public submitReport() {
    
    let formDigest = "";
    let reportedby: string = this.state.userEmail;
    let page: string = this.state.location;
    let note: string = this.state.feedback;
    axios.post('https://peoplesmortgagecompany.sharepoint.com/_api/contextinfo')
    .then((res) => {
      console.log(res);
      formDigest = res.data.FormDigestValue;
      console.log('formdigest: ' + formDigest);
    })
    .then(() => {
      axios({
        method: 'POST',
        url: "https://peoplesmortgagecompany.sharepoint.com/sites/intranet/loanhelp/_api/web/lists/GetByTitle('User%20Feedback')/items",
        headers: {
          "X-RequestDigest": formDigest,
          "Accept": "application/json;odata=verbose",
          "content-type": "application/json;odata=verbose",
        },
        data: {
          '__metadata': {
            'type': 'SP.Data.User_x0020_FeedbackListItem'
          },
          'Title': new Date(),
          'ReportedBy': reportedby,
          'Page': page,
          'Note': note
        }
      })
      .then((res) => {
        console.log(res);
      })
      .catch((err) => {
        console.log(err);
      })
    })
  }

  public onClosePanel() {
    this.setState({
      buttonVal: false
    })
  }

  public buttonClicked() {
    this.setState({
      buttonVal: true
    });
  }
  
  componentDidMount() {
    axios({
      method:'GET',
      url:'https://peoplesmortgagecompany.sharepoint.com/_api/web/CurrentUser',
    })
    .then((res) => {
      this.setState({
        location: window.location.href,
        userEmail: res.data.Email
      });
    });
  }
}