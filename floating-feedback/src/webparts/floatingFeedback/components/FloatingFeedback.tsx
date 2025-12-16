import * as React from 'react';
import styles from './FloatingFeedback.module.scss';
import { IFloatingFeedbackProps } from './IFloatingFeedbackProps';
import {
  PrimaryButton,
  DefaultButton,
  Dialog,
  DialogType,
  DialogFooter,
  TextField,
  Rating,
  MessageBar,
  MessageBarType
} from '@fluentui/react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IFloatingFeedbackState {
  isModalOpen: boolean;
  title: string;
  description: string;
  rating: number;
  selectedCategories: string[]; // Selected category values
  availableCategories: string[]; // All available choices from SP
  allowMultipleValues: boolean; // Does the SP column allow multi-select?
  isSubmitting: boolean;
  message: string;
  messageType: MessageBarType;
}

export default class FloatingFeedback extends React.Component<IFloatingFeedbackProps, IFloatingFeedbackState> {

  constructor(props: IFloatingFeedbackProps) {
    super(props);
    this.state = {
      isModalOpen: false,
      title: '',
      description: '',
      rating: 3,
      selectedCategories: [],
      availableCategories: [],
      allowMultipleValues: true, // Default to true, updated on fetch
      isSubmitting: false,
      message: '',
      messageType: MessageBarType.info
    };
  }

  public componentDidMount(): void {
    this._getCategoryChoices();
  }

  public componentDidUpdate(prevProps: IFloatingFeedbackProps): void {
    if (this.props.listName !== prevProps.listName) {
      this._getCategoryChoices();
    }
  }

  private _getCategoryChoices(): void {
    const { listName, spHttpClient, siteUrl } = this.props;
    if (!listName) return;

    // Hardcoded to 'Category' column as per requirement
    spHttpClient.get(`${siteUrl}/_api/web/lists/getByTitle('${listName}')/fields/getByInternalNameOrTitle('Category')?$select=Choices,AllowMultipleValues,TypeAsString`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .then((data) => {
        if (data.Choices) {
          this.setState({
            availableCategories: data.Choices,
            allowMultipleValues: !!data.AllowMultipleValues || data.TypeAsString === 'MultiChoice'
          });
        }
      })
      .catch(err => console.error('Error fetching category choices', err));
  }

  private _toggleCategory = (category: string): void => {
    const { selectedCategories, allowMultipleValues } = this.state;
    const index = selectedCategories.indexOf(category);
    let newSelection: string[];

    if (index > -1) {
      // Remove
      newSelection = selectedCategories.filter(c => c !== category);
    } else {
      // Add
      if (allowMultipleValues) {
        newSelection = [...selectedCategories, category];
      } else {
        // If single select, replace the selection
        newSelection = [category];
      }
    }

    this.setState({ selectedCategories: newSelection });
  }

  public render(): React.ReactElement<IFloatingFeedbackProps> {
    const { position } = this.props;
    const { isModalOpen, title, description, rating, selectedCategories, availableCategories, allowMultipleValues, isSubmitting, message, messageType } = this.state;


    // Calculate style based on position
    const btnStyle: React.CSSProperties = position === 'Top' ? { top: '20px' } : { bottom: '20px' };

    return (
      <div className={styles.floatingFeedback}>
        <div className={styles.floatingBtn} style={btnStyle}>
          <PrimaryButton
            text="Feedback"
            iconProps={{ iconName: 'Feedback' }}
            onClick={this._openModal}
          />
        </div>

        <Dialog
          hidden={!isModalOpen}
          onDismiss={this._closeModal}
          dialogContentProps={{
            type: DialogType.close,
            title: 'Submit Feedback',
            subText: 'We would love to hear your thoughts.'
          }}
          modalProps={{
            isBlocking: false,
            styles: { main: { maxWidth: 900, minWidth: 600 } }
          }}
        >
          <div className={styles.modernFeedbackForm}>
            <TextField
              label="Title"
              value={title}
              onChange={(e, val) => this.setState({ title: val || '' })}
              disabled={isSubmitting}
              required
            />
            <TextField
              label="Description"
              multiline
              rows={4}
              value={description}
              onChange={(e, val) => this.setState({ description: val || '' })}
              disabled={isSubmitting}
              required
            />

            <div style={{ marginBottom: 15 }}>
              <label className="ms-Label">Category</label>
              <div className={styles.categoryContainer}>
                {availableCategories.map((cat) => (
                  <div
                    key={cat}
                    className={`${styles.categoryButton} ${selectedCategories.indexOf(cat) > -1 ? styles.categoryButtonSelected : ''}`}
                    onClick={() => !isSubmitting && this._toggleCategory(cat)}
                  >
                    {cat}
                  </div>
                ))}
                {availableCategories.length === 0 && <span>No categories found in the 'Category' column.</span>}
              </div>
            </div>

            <div className={styles.ratingContainer}>
              <label className="ms-Label">Rating</label>
              <Rating
                min={1}
                max={5}
                rating={rating}
                onChange={(e, val) => this.setState({ rating: val || 0 })}
                disabled={isSubmitting}
              />
            </div>

            {message && (
              <div className={styles.messageBar}>
                <MessageBar messageBarType={messageType}>
                  {message}
                </MessageBar>
              </div>
            )}
          </div>
          <DialogFooter>
            <PrimaryButton onClick={this._submitFeedback} text="Submit" disabled={isSubmitting || !title || !description} />
            <DefaultButton onClick={this._closeModal} text="Cancel" disabled={isSubmitting} />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }

  private _openModal = (): void => {
    this.setState({
      isModalOpen: true,
      message: '',
      messageType: MessageBarType.info,
      title: '',
      description: '',
      rating: 3,
      selectedCategories: []
    });
  }

  private _closeModal = (): void => {
    this.setState({ isModalOpen: false });
  }



  private _submitFeedback = (): void => {
    const { listName, spHttpClient, siteUrl, userDisplayName } = this.props;
    const { title, description, rating, selectedCategories, allowMultipleValues } = this.state;

    if (!listName) {
      this.setState({ message: 'Error: No list configured.', messageType: MessageBarType.error });
      return;
    }

    this.setState({ isSubmitting: true, message: '' });

    const requestBody: any = {}; // eslint-disable-line @typescript-eslint/no-explicit-any

    // User requested mappings:
    // Title -> Title
    // Description -> FeedbackText
    // Category -> Category
    // SubmittedBy -> userDisplayName (Text)

    requestBody.Title = title;
    requestBody.FeedbackText = description;
    requestBody.SubmittedBy = userDisplayName;
    requestBody.Rating = rating;
    requestBody.PageName = this.props.pageName;

    // Multi-choice handling for Category
    if (selectedCategories.length > 0) {
      if (allowMultipleValues) {
        // For odata=nometadata, send array directly
        requestBody.Category = selectedCategories;
      } else {
        // Single value for single-choice column
        requestBody.Category = selectedCategories[0];
      }
    }

    spHttpClient.post(`${siteUrl}/_api/web/lists/getByTitle('${listName}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: JSON.stringify(requestBody)
      })
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          this.setState({
            isSubmitting: false,
            message: 'Feedback submitted successfully!',
            messageType: MessageBarType.success,
            title: '',
            description: '',
            rating: 3,
            selectedCategories: []
          });
          setTimeout(() => this._closeModal(), 2000);
        } else {
          return response.json().then(error => {
            this.setState({
              isSubmitting: false,
              message: `Error: ${error.error?.message?.value || response.statusText}`,
              messageType: MessageBarType.error
            });
          });
        }
      })
      .catch((error) => {
        this.setState({
          isSubmitting: false,
          message: `Error: ${error.message}`,
          messageType: MessageBarType.error
        });
      });
  }
}
