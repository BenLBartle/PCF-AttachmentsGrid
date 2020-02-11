import * as React from 'react';
import { Card, ICardTokens, ICardSectionStyles, ICardSectionTokens } from '@uifabric/react-cards';
import { FontWeights } from '@uifabric/styling';
import {
  Icon,
  IIconStyles,
  Image,
  Persona,
  Stack,
  IStackTokens,
  Text,
  ITextStyles,
  PersonaSize
} from 'office-ui-fabric-react';
import { format } from 'timeago.js';
import filesize from 'filesize';
// Helper imports to generate data for this particular examples. Not exported by any package.

const stackTokens: IStackTokens = { childrenGap: 10 };

const alertClicked = (): void => {
  alert('Clicked');
};

const siteTextStyles: ITextStyles = {
  root: {
    color: '#025F52',
    fontWeight: FontWeights.semibold
  }
};
const descriptionTextStyles: ITextStyles = {
  root: {
    color: '#333333',
    fontWeight: FontWeights.semibold
  }
};
const helpfulTextStyles: ITextStyles = {
  root: {
    color: '#333333',
    fontWeight: FontWeights.regular
  }
};
const iconStyles: IIconStyles = {
  root: {
    color: '#0078D4',
    fontSize: 16,
    fontWeight: FontWeights.regular
  }
};
const footerCardSectionStyles: ICardSectionStyles = {
  root: {
    borderTop: '1px solid #F3F2F1'
  }
};

const cardTokens: ICardTokens = { childrenMargin: 12 };
const footerCardSectionTokens: ICardSectionTokens = { padding: '12px 0px 0px' };



export interface IXrmAttachmentControlProps {
  defaultAttachments: IXrmAttachmentProps[]
  // webApi: ComponentFramework.WebApi
  // entityId: string
  // entityLogicalName: string
};

export interface IXrmAttachmentProps {
  id: string;
  name: string;
  icon: string;
  filename: string;
  fileextension: string;
  size: number;
  createdBy: string;
  createdOn: Date;
};

export interface IXrmAttachmentControlState {
  attachments: IXrmAttachmentProps[];
}

export class XrmAttachmentControl extends React.Component<IXrmAttachmentControlProps, IXrmAttachmentControlState> {
  constructor(props: IXrmAttachmentControlProps) {
    super(props);
    this.state = {
      attachments: props.defaultAttachments
    };
  }

  public render(): JSX.Element {

    let list = this.state.attachments.map(attachment => {
      return (
        <Card tokens={cardTokens}>
          <Card.Section>
            <Icon iconName={attachment.icon} styles={iconStyles} />
            <Text variant="small" styles={siteTextStyles}>
              {attachment.filename}.{attachment.fileextension}
            </Text>
            <Text styles={descriptionTextStyles}>{attachment.name}</Text>
          </Card.Section>
          <Card.Item>
            <Persona size={PersonaSize.size40} text={attachment.createdBy} secondaryText={format(attachment.createdOn)} />
          </Card.Item>
          <Card.Section horizontal styles={footerCardSectionStyles} tokens={footerCardSectionTokens}>
            <Icon iconName="Delete" styles={iconStyles} onClick={alertClicked} />
            <Stack.Item grow={1}>
              <span />
            </Stack.Item>
            <Text variant="small" styles={helpfulTextStyles}>
              {filesize(attachment.size)}
            </Text>
          </Card.Section>
        </Card>
      );
    });
    return (<Stack horizontal wrap tokens={stackTokens}>{list}</Stack>);
  }

}