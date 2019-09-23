import * as React from 'react';
import TimeAgo from 'react-timeago'
import { Card, ICardTokens, ICardSectionStyles, ICardSectionTokens } from '@uifabric/react-cards';
import { FontWeights } from '@uifabric/styling';
import {
  ActionButton,
  IButtonStyles,
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
const backgroundImageCardSectionStyles: ICardSectionStyles = {
  root: {
    backgroundImage: 'url(https://placehold.it/256x144)',
    backgroundPosition: 'center center',
    backgroundSize: 'cover',
    height: 144
  }
};
const dateTextStyles: ITextStyles = {
  root: {
    color: '#505050',
    fontWeight: 600
  }
};
const subduedTextStyles: ITextStyles = {
  root: {
    color: '#666666'
  }
};
const actionButtonStyles: IButtonStyles = {
  root: {
    border: 'none',
    color: '#333333',
    height: 'auto',
    minHeight: 0,
    minWidth: 0,
    padding: 0,

    selectors: {
      ':hover': {
        color: '#0078D4'
      }
    }
  },
  textContainer: {
    fontSize: 12,
    fontWeight: FontWeights.semibold
  }
};

const sectionStackTokens: IStackTokens = { childrenGap: 30 };
const cardTokens: ICardTokens = { childrenMargin: 12 };
const footerCardSectionTokens: ICardSectionTokens = { padding: '12px 0px 0px' };
const backgroundImageCardSectionTokens: ICardSectionTokens = { padding: 12 };
const agendaCardSectionTokens: ICardSectionTokens = { childrenGap: 0 };
const attendantsCardSectionTokens: ICardSectionTokens = { childrenGap: 6 };



export interface IXrmAttachmentControlProps {
  defaultAttachments: IXrmAttachmentProps[]
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
        <Card onClick={alertClicked} tokens={cardTokens}>
          <Card.Item fill>
            <Image src="https://placehold.it/256x144" width="100%" alt="Placeholder image." />
          </Card.Item>
          <Card.Section>
            <Text variant="small" styles={siteTextStyles}>
              {attachment.filename}.{attachment.fileextension}
            </Text>
            <Text styles={descriptionTextStyles}>{attachment.name}</Text>
          </Card.Section>
          <Card.Item>
      <Persona size={PersonaSize.size24} text={attachment.createdBy} secondaryText={<TimeAgo>attachment.CreateOn</TimeAgo>} />
          </Card.Item>
          <Card.Section horizontal styles={footerCardSectionStyles} tokens={footerCardSectionTokens}>
            <Icon iconName="Edit" styles={iconStyles} />
            <Icon iconName="Delete" styles={iconStyles} />
            <Stack.Item grow={1}>
              <span />
            </Stack.Item>
            <Text variant="small" styles={helpfulTextStyles}>
              1024 KB
          </Text>
          </Card.Section>
        </Card>
      );
    });
    return (<Stack horizontal wrap tokens={stackTokens}>{list}</Stack>);
  }

}