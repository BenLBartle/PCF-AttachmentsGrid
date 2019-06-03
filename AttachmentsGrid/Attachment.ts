import { EntityReference } from "./EntityReference";
export class Attachment {
	attachmentId: EntityReference;
	name: string;
	extension: string;
	entityType: string;
	constructor(attachmentId: EntityReference, name: string, extension: string) {
		this.attachmentId = attachmentId;
		this.name = name;
		this.extension = extension;
	}
}
