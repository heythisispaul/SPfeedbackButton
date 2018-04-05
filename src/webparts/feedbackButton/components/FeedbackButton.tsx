import * as React from 'react';
import styles from './FeedbackButton.module.scss';
import { IFeedbackButtonProps } from './IFeedbackButtonProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IFeedbackButtonState } from './IFeedbackButtonState';
import { DefaultButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import axios from 'axios';

export default class FeedbackButton extends React.Component<IFeedbackButtonProps, IFeedbackButtonState> {

  constructor(props) {
    super(props);
    this.buttonClicked = this.buttonClicked.bind(this);
    this.onClosePanel = this.onClosePanel.bind(this);
    this.submitReport = this.submitReport.bind(this);
    this.helpfulOption = this.helpfulOption.bind(this);
    this.trainingOption = this.trainingOption.bind(this);
    this.state = {
      location: "",
      feedback: "",
      buttonVal: false,
      userEmail: "",
      submitStatus: false,
      helpfulSelection: "",
      moreTraining: "",
      helpfulButton: "",
      trainingButton: ""
    }
  }

  public render(): React.ReactElement<IFeedbackButtonProps> {
    let disabled:any = this.state.buttonVal ? 'disabled' : null;
    return (
          <div>
              <DefaultButton
              disabled={false}
              primary={true}
              text={this.props.buttonText}
              onClick={ this.buttonClicked }
              />
              { this.state.buttonVal ? 
                 <Panel
                 isOpen={ this.state.buttonVal }
                 type={ PanelType.smallFixedFar }
                 onDismiss={ this.onClosePanel }
                 headerText='Your LoanHelp Experience'
                 closeButtonAriaLabel='Close'>
                <div className={styles.submitButton} style={{marginTop: 10 + 'px'}}>
                  {!this.state.submitStatus ? 
                  <div>
                    <ChoiceGroup
                      defaultSelectedKey='B'
                      selectedKey={this.state.helpfulButton}
                      options={ [
                        {
                          key: 'A',
                          text: 'Yes',
                          disabled: this.state.submitStatus ? true : false
                        },
                        {
                          key: 'B',
                          text: 'No',
                          disabled: this.state.submitStatus ? true : false
                        }
                      ] }
                      onChange={ this.helpfulOption }
                      label='Was this page helpful?'
                      required={ true }
                    />
                    <ChoiceGroup
                      defaultSelectedKey='B'
                      selectedKey={this.state.trainingButton}
                      options={ [
                        {
                          key: 'A',
                          text: 'Yes',
                          disabled: this.state.submitStatus ? true : false
                        },
                        {
                          key: 'B',
                          text: 'No',
                          disabled: this.state.submitStatus ? true : false
                        }
                      ] }
                      onChange={ this.trainingOption }
                      label='Was this page helpful?'
                      required={ true }
                    />
                    <TextField 
                      label="Comments:" 
                      value={this.state.feedback} 
                      multiline rows={5} onChanged={(newVal:any) => { this.setState({feedback: newVal})}}
                      disabled={this.state.submitStatus ? true : false}
                    />
                    <div className={styles.submitButton} style={{marginTop: 10 + 'px'}}>
                      <DefaultButton primary={true} text='Submit' onClick={this.submitReport} disabled={this.state.submitStatus ? true : false}/>
                    </div>
                  </div>
                  : <div><b>Thank you for your Feedback</b></div>}  
                </div>
               </Panel>
              : null}
          </div>
    );
  }

  private helpfulOption(ev: React.FormEvent<HTMLInputElement>, option: any): void {
    this.setState({
      helpfulButton: option.key,
      helpfulSelection: option.text
    })
  }

  private trainingOption(ev: React.FormEvent<HTMLInputElement>, option: any): void {
    this.setState({
      trainingButton: option.key,
      moreTraining: option.text
    })
  }

  public submitReport() {
    let formDigest:any = "";
    let reportedby: string = this.state.userEmail;
    let page: string = this.state.location;
    let note: string = this.state.feedback;
    let training = this.state.moreTraining;
    let helpful = this.state.helpfulSelection;
    axios.post('https://peoplesmortgagecompany.sharepoint.com/sites/intranet/loanhelp/_api/contextinfo')
    .then((res) => {
      formDigest = res.data.FormDigestValue;
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
          'Note': note,
          'MoreTraining': training,
          'PageIsHelpful': helpful
        }
      })
      .then((res) => {
        this.setState({
          submitStatus: true,
          feedback: ""
        })
        console.log(res);
      })
      .catch((err) => {
        console.log(err);
      })
    })
  }

  public onClosePanel() {
    this.setState({
      buttonVal: false,
      submitStatus: false
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