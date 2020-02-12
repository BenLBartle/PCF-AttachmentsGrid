import { IInputs, IOutputs } from "./generated/ManifestTypes";
import 'bootstrap';
import { DropHandler } from "./drophandler/drophandler";
import { Attachment } from "./Attachment";
import { EntityReference } from "./EntityReference";
import { Subject } from "rxjs";
import { IXrmAttachmentControlProps } from "./IXrmAttachmentControlProps";
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { initializeIcons } from '@uifabric/icons';
import { XrmAttachmentControl } from "./AttachmentsGrid";
initializeIcons();

class AttachmentRef {
	id: string;
	type: string;
	constructor(id: string, type: string) {
		this.id = id;
		this.type = type;
	}
}

class FileToDownload implements ComponentFramework.FileObject {
	fileContent: string;
	fileName: string;
	fileSize: number;
	mimeType: string;
}

export class AttachmentsGrid implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	// Reference to the notifyOutputChanged method
	private notifyOutputChanged: () => void;
	// Reference to the container div
	private theContainer: HTMLDivElement;

	private props: IXrmAttachmentControlProps;
	private _reference: EntityReference;

	private _context: ComponentFramework.Context<IInputs>;
	private _container: HTMLDivElement;
	private _divControl: HTMLDivElement;

	// Define Input Elements
	public _dropElement: HTMLDivElement;

	public _progressElement: HTMLDivElement;
	public _progressBar: HTMLDivElement;

	private _dropHandler: DropHandler;

	private _apiClient: ComponentFramework.WebApi;

	private _attachmentSource = new Subject<Attachment>();

	private _attachmentContainer: HTMLDivElement;

	attachmentAdded$ = this._attachmentSource.asObservable();

	/**
	 * Empty constructor.
	 */
	constructor() {

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='starndard', it will receive an empty div element within which it can render its content.
	 */
	public async init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {

		this.theContainer = container;

		this._reference = new EntityReference(
			(<any>context).page.entityTypeName,
			(<any>context).page.entityId
		)

		this._reference = new EntityReference("contact", "2");
	}


	private onClickDelete(divCard: HTMLDivElement, attachment: AttachmentRef) {
		//show confirm or cancel action
		let confirmDelete: HTMLDivElement = document.createElement("div");
		confirmDelete.className = "confirmDelete";
		divCard.appendChild(confirmDelete);
		//add html button group
		confirmDelete.innerHTML = `<div class='btn-group btn-group-sm' role='group'><button id='deleteButton' type='button' class='btn btn-secondary'>Delete</button><button id='cancelButton' type='button' class='btn btn-secondary'>Cancel</button></div>`;

		//get button elements from html
		let cancelButton = document.getElementById("cancelButton");
		cancelButton = <HTMLButtonElement>cancelButton;
		let deleteButton = document.getElementById("deleteButton");
		deleteButton = <HTMLButtonElement>deleteButton;

		if (cancelButton) {
			cancelButton.addEventListener("click", this.onClickCancelDelete.bind(this, confirmDelete));
		}
		if (deleteButton) {
			deleteButton.addEventListener("click", this.onClickConfirmDelete.bind(this, attachment));
		}

	}

	private onClickConfirmDelete(attachment: AttachmentRef) {
		//get the attachment id
		let id = attachment.id;

		//delete the attachment
		this._context.webAPI.deleteRecord("annotation", id).then(
			function success(result) {
				console.log("Successfully deleted the record")
				return result;
			}
		).then(result => {
			console.log(result.id)
		})
	}

	private onClickCancelDelete(confirmDelete: HTMLDivElement) {
		//clear the buttons
		confirmDelete.innerHTML = "";
	}

	private findImage(img: HTMLImageElement, item: Attachment) {
		//find the image
		this._context.resources.getResource(`${item.extension.toLowerCase()}.png`,
			content => {
				this.setImage(img, "png", content);
			},
			() => {
				this.showError();
			});
	}

	private setImage(element: HTMLImageElement, fileType: string, fileContent: string): void {
		//set the image to the img element
		fileType = fileType.toLowerCase();
		let imageUrl: string = `data:image/${fileType};base64, ${fileContent}`;
		element.src = imageUrl;
	}

	private showError(): void {
		console.log('error while downloading .png');
	}

	private async getAttachments(reference: EntityReference): Promise<Attachment[]> {
		//webapi query to find any attachments
		let query = `?$select=annotationid,filename,filesize,createdon,mimetype&$filter=filename ne null and _objectid_value eq ${reference.id} and objecttypecode eq '${reference.typeName}' &$orderby=createdon desc`;
		try {
			const result = await this._context.webAPI.retrieveMultipleRecords("annotation", query);
			let items: Attachment[] = [];
			for (let i = 0; i < result.entities.length; i++) {
				let ent = result.entities[i];
				let item = new Attachment(new EntityReference("annotation", ent["annotationid"].toString()), ent["filename"].split('.')[0], ent["filename"].split('.')[1].toLowerCase(), false);

				this._attachmentSource.next(item);
			}
			return items;
		}
		catch (error) {
			console.log(error.message);
			let items_1: Attachment[] = [];
			return items_1;
		}
	}

	private getAttachment(id: string, type: string): Promise<FileToDownload> {
		//query the attachement id
		return this._context.webAPI.retrieveRecord("annotation", id).then(
			function success(result) {
				let file: FileToDownload = new FileToDownload();
				file.fileContent =
					type == "annotation" ? result["documentbody"] : result["body"];
				file.fileName = result["filename"];
				file.fileSize = result["filesize"];
				file.mimeType = result["mimetype"];
				return file;

			}

		);
	}

	private onClickAttachment(attachment: AttachmentRef): void {
		//get the attachment id
		let id = attachment.id;
		let type = attachment.type;

		this.getAttachment(id, type).then((r: any) => {
			this._context.navigation.openFile(r);
		}
		);
	}


	public updateView(context: ComponentFramework.Context<IInputs>): void {

		const attachments = [{
			id: "1",
			name: "How you doin'?",
			icon: "WordDocument",
			filename: "Pickup Lines",
			fileextension: ".docx",
			size: 272828182,
			createdBy: "Joey Tribiani",
			createdOn: new Date(1974, 10)
		},
		{
			id: "1",
			name: "Secret Map",
			icon: "PDF",
			filename: "Sonics Layer",
			fileextension: "pdf",
			size: 28930302,
			createdBy: "Dr Robotnik",
			createdOn: new Date(2020, 1)
		},
		{
			id: "1",
			name: "Evil Plan",
			icon: "PowerPointDocument",
			filename: "World Domination",
			fileextension: "ppt",
			size: 4000,
			createdBy: "Ben Bartle",
			createdOn: new Date(2019, 8)
		},
		{
			id: "1",
			name: "No one will know it's me",
			icon: "PDF",
			filename: "Disguise",
			fileextension: "pdf",
			size: 293933,
			createdBy: "Sparticus",
			createdOn: new Date(2009, 12)
		},
		{
			id: "1",
			name: "Damnit again",
			icon: "PDF",
			filename: "Speeding Ticket",
			fileextension: "pdf",
			size: 1220330,
			createdBy: "Chuck Jaeger",
			createdOn: new Date()
		},
		{
			id: "1",
			name: "Muhaa haa haa",
			icon: "ExcelDocument",
			filename: "Enemies List",
			fileextension: "xlsx",
			size: 12,
			createdBy: "Pol Pot",
			createdOn: new Date(2020, 1)
		}];

		this.props = {
			defaultAttachments: attachments,
			webApi: context.webAPI,
			entityReference: this._reference
		}

		// Render the React component into the div container
		ReactDOM.render(
			// Create the React component
			React.createElement(
				XrmAttachmentControl,
				this.props
			),
			this.theContainer
		);
	}

	public getOutputs(): IOutputs {
		return {};
	}

	public destroy(): void {
		// Add code to cleanup control if necessary

	}

}