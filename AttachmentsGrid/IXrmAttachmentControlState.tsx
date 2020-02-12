import { IXrmAttachmentProps } from "./IXrmAttachmentProps";
export interface IXrmAttachmentControlState {
    attachments: IXrmAttachmentProps[];
    progressShow: boolean;
    progressTotalFilesToUpload: number;
    progressCurrentFilesUploaded: number;
}
