import { EntityReference } from "./EntityReference";
export class Attachment {
	attachmentId: EntityReference;
	name: string;
	extension: string;
	entityType: string;
	deleted: boolean;
	constructor(attachmentId: EntityReference, name: string, extension: string, deleted: boolean) {
		this.attachmentId = attachmentId;
		this.name = name;
		this.extension = extension;
		this.deleted = deleted;
	}
}
