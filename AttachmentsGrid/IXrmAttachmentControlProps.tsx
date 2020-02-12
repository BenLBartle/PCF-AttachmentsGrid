import { IXrmAttachmentProps } from './IXrmAttachmentProps';
import { DropHandler } from './drophandler/drophandler';
import { EntityReference } from './EntityReference';
export interface IXrmAttachmentControlProps {
  defaultAttachments: IXrmAttachmentProps[];
  entityReference: EntityReference,
  webApi: ComponentFramework.WebApi
};
