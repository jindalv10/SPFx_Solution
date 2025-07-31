import * as React from 'react';
import styles from './FileUploader.module.scss';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { AllRoles, IAttachment } from 'pd-nomination-library';

type FileUploaderProps = {
  onFilesChanged(files: any[]);
  docType?: string;
  context: any;
  disabled: boolean;
  files?: IAttachment[],
  role: string
  actor?:string;
  onDocumentDelete?: (fileName: string) => void;
};
type FileUploaderState = {
  dragging: boolean;
  files: IAttachment[] | null;
  className: string;
  filePickerResult?: any;
};

export class FileUploader extends React.Component<FileUploaderProps, FileUploaderState> {
  private fileUploaderInput: HTMLElement | null = null;
  private dragEventCounter = 0;

  constructor(props: FileUploaderProps) {
    super(props);
    this.state = {
      dragging: false,
      files: null,
      className: styles.file_uploader
    };
  }

  private dragenterListener = (event: React.DragEvent<HTMLDivElement>) => {
    this.overrideEventDefaults(event);
    this.dragEventCounter++;
    if (event.dataTransfer.items && event.dataTransfer.items[0]) {
      this.setState({ dragging: true, className: styles.file_uploader__dragging });
    } else if (
      event.dataTransfer.types &&
      event.dataTransfer.types[0] === "Files"
    ) {
      // This block handles support for IE - if you're not worried about
      // that, you can omit this
      this.setState({ dragging: true, className: styles.file_uploader__dragging });
    }
  }

  private dragleaveListener = (event: React.DragEvent<HTMLDivElement>) => {
    this.overrideEventDefaults(event);
    this.dragEventCounter--;

    if (this.dragEventCounter === 0) {
      this.setState({ dragging: false, className: "" });
    }
  }

  private dropListener = (event: React.DragEvent<HTMLDivElement>) => {
    this.overrideEventDefaults(event);
    this.dragEventCounter = 0;
    this.setState({ dragging: false, className: "" });

    if (event.dataTransfer.files && event.dataTransfer.files[0]) {
      this.refreshState(event.dataTransfer.files);
    }
  }

  private overrideEventDefaults = (event: Event | React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    event.stopPropagation();
  }

  private onSelectFileClick = () => {
    if (this.fileUploaderInput)
      this.fileUploaderInput.click();
  }

  private onFileChanged = (event: React.ChangeEvent<HTMLInputElement>) => {
    if (event.target.files && event.target.files[0]) {
      this.refreshState(event.target.files);
    }
    event.currentTarget.value = null;
  }

  private refreshState = (sourceFiles: FileList) => {
    var files = this.state.files;
    if (files == null) {
      files = [];
    }
    for (var i = 0; i < sourceFiles.length; i++) {
      if (files && files.filter((f: any) => {
        var fileName = f.name || f.fileName;
        return decodeURIComponent(fileName).toLowerCase() === sourceFiles[i].name.toLowerCase();
      }).length === 0) {
        files.push({ id: 0, file: sourceFiles[i], attachmentType: this.props.docType, attachmentName: sourceFiles[0].name, attachmentBy: this.props.role });
      }
    }
    this.setState({ files: files });
    this.onFilesChanged();
  }

  private onFilesChanged = () => {
    var isMetadataValid = true;
    setTimeout(() => {
      this.props.onFilesChanged(this.state.files);
    }, 500);
  }

  private onDeleteDoc = (file: IAttachment) => {
    var files = this.state.files;
    var selectedIndex: number = -1;
    var filteredFiles = files.filter((fi: any, index) => {
      var fileName = fi.attachmentName || fi.attachmentName;
      if (fileName === (file.attachmentName || file.attachmentName))
      {
        selectedIndex = index;

        if(this.props.role == AllRoles.QC || this.props.role == AllRoles.NOMINATOR)
        {
          this.props.onDocumentDelete(fileName);
          files = null;
        }  
      }
      else
        return fi;
    });

    this.setState({
      files: filteredFiles
    });
    this.onFilesChanged();
  }

  public componentDidMount() {
    const { files } = this.props;
    if (files) {
      this.setState({ files: files });
    }
    window.addEventListener("dragover", (event: Event) => {
      this.overrideEventDefaults(event);
    });
    window.addEventListener("drop", (event: Event) => {
      this.overrideEventDefaults(event);
    });
  }

  public componentWillUnmount() {
    window.removeEventListener("dragover", this.overrideEventDefaults);
    window.removeEventListener("drop", this.overrideEventDefaults);
  }


  public render() {
    return (
      <div>
        {
          !this.props.disabled && <div
            className={`${styles.file_uploader}`}
            onDrag={this.overrideEventDefaults}
            onDragStart={this.overrideEventDefaults}
            onDragEnd={this.overrideEventDefaults}
            onDragOver={this.overrideEventDefaults}
            onDragEnter={this.dragenterListener}
            onDragLeave={this.dragleaveListener}
            onDrop={this.dropListener}
          >
            <span>Drag & Drop Files</span>
            <span>or</span>
            <span onClick={this.onSelectFileClick} className={styles.file_uploader_select}>
              Select files from your computer
            </span>

          </div>
        }
        <div className={styles.file_uploader__contents}>
          {
            this.state.files && this.state.files.length > 0 &&
            <table className={styles.attachmentsTable}>
              <tr>
                <th>Name</th>
                {/* <th>Document Type</th> */}
                {!this.props.disabled && <th>Remove</th>}
              </tr>
              {
                this.state.files.map((f: IAttachment, index) => {
                  return <tr>
                    <td><a href={f.attachmentUrl} target="_blank">{(f.attachmentName || f.attachmentName)}</a></td>
                    {/* <td>{(f.attachmentType || f.attachmentType)}</td>                                     */}
                    {!this.props.disabled && <td>
                      <span onClick={() => this.onDeleteDoc(f)}>
                        <Icon iconName="Delete" className={styles.msIcon} />
                      </span>
                    </td>}
                  </tr>;
                })
              }
            </table>
            // : <span className={styles.deliverableMessage}>No Files Selected.  Use the above control to upload deliverables.</span>
          }
        </div>
        <input
          ref={el => (this.fileUploaderInput = el)}
          type="file" multiple={true}
          disabled={this.props.disabled}
          className={styles.file_uploader__input}
          onChange={this.onFileChanged}
        />
      </div>
    );
  }
}