import { EntityReference } from "./EntityReference";
export class Attachment {
	attachmentId: EntityReference;
	name: string;
	extension: string;
	entityType: string;
	deleted: boolean;
	size: number;
	constructor(attachmentId: EntityReference, name: string, extension: string, deleted: boolean, size: number) {
		this.attachmentId = attachmentId;
		this.name = name;
		this.extension = extension;
		this.deleted = deleted;
		this.size = size;
	}
}
