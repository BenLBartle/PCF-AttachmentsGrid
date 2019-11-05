import { IInputs, IOutputs } from "./generated/ManifestTypes";
import 'bootstrap';
import { DropHandler } from "./drophandler/drophandler";
import { Attachment } from "./Attachment";
import { EntityReference } from "./EntityReference";
import { Subject } from "rxjs";
import { library, dom } from '@fortawesome/fontawesome-svg-core'
import { faDownload } from '@fortawesome/free-solid-svg-icons'
import { faTrashAlt, faFilePdf, faFileAlt, faFileArchive, faFileExcel, faFileImage, faFilePowerpoint, faFileWord, faFileCode } from '@fortawesome/free-regular-svg-icons'

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

class OpenFileOptions implements ComponentFramework.NavigationApi.OpenFileOptions {
	openMode: ComponentFramework.NavigationApi.Types.OpenFileMode;
}

export class AttachmentsGrid implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private _context: ComponentFramework.Context<IInputs>;
	private _apiClient: ComponentFramework.WebApi;

	// Define Input Elements
	private _refreshFileNameAfterUpload: boolean;

	// Define Standard container element
	private _container: HTMLDivElement;

	// Define Display Elements
	public _dropElement: HTMLDivElement;
	public _progressElement: HTMLDivElement;
	public _progressBar: HTMLDivElement;
	public _attachmentsNotAvialbleElement: HTMLDivElement;
	private _attachmentContainer: HTMLDivElement;
	
	// Define Drop Handler
	private _dropHandler: DropHandler;

	// Define attachment grid data
	private _attachmentSource = new Subject<Attachment>();
	attachmentAdded$ = this._attachmentSource.asObservable();

	// Define entity data
	private _reference: EntityReference;
	private _initialized: boolean;

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

		// Add control initialization code
		this._context = context;
		this._container = document.createElement("div");
		this._container.classList.add("bootstrap-iso-attachments-grid");
		this._apiClient = context.webAPI;
		this._refreshFileNameAfterUpload = (context.parameters.RefreshAttachmentNameAfterUpload && context.parameters.RefreshAttachmentNameAfterUpload.raw && context.parameters.RefreshAttachmentNameAfterUpload.raw.toLowerCase() === 'true') ? true : false;

		// Add fontawesome icons
		library.add(faDownload, faTrashAlt, faFilePdf, faFileAlt, faFileArchive, faFileExcel, faFileImage, faFilePowerpoint, faFileWord, faFileCode);

		// Add dom watcher for fontawesome.
		// In order to use this update node_modules\@fortawesome\fontawesome-svg-core\index.d.ts file.
		// A pull request will be put in with the fontawesome team to fix this.
		// 
		//Update the watch line in the following code
		//export interface DOM {
		//   i2svg(params?: { node: Node; callback: () => void }): Promise<void>;
		//   css(): string;
		//   insertCss(): string;
		//   watch(): void;
		//}
		//
		//to =>
		//
		//	watch(params?: { autoReplaceSvgRoot: Node; observeMutationsRoot: Node}): void;
		//
		dom.watch({
			autoReplaceSvgRoot: container,
			observeMutationsRoot: container
		})

		// Progress Bar Elements
		this._progressElement = document.createElement("div");
		this._progressElement.classList.add("progress");
		this._progressElement.style.height = "5px";
		this._progressElement.style.visibility = "hidden";

		this._progressBar = document.createElement("div");
		this._progressBar.classList.add("progress-bar");
		this._progressBar.style.width = "0%";

		// Layout Elements
		this._dropElement = document.createElement("div");
		this._dropElement.classList.add("drop-zone");

		// Attachment Elements
		this._attachmentContainer = document.createElement("div");
		this._attachmentContainer.className = "card-columns";
		this._dropElement.appendChild(this._attachmentContainer);

		this._attachmentsNotAvialbleElement =  document.createElement("div");
		this._attachmentsNotAvialbleElement.innerText = "This record hasn't been created yet.  To add or view attachments, save it first.";
		this._attachmentsNotAvialbleElement.classList.add("attachmentsNotAvailable", "hidden");
		this._progressElement.append(this._progressBar);
		this._container.append(this._attachmentsNotAvialbleElement, this._progressElement, this._dropElement);

		// Bind to parent container
		container.append(this._container);		

		this.attachmentAdded$.subscribe(a => {
			this.createBSCard(a);
		})

		//await this.initAttachmentsAndDropHandler(this._context);
	}

	private createBSCard(attachment: Attachment) {

		let divCard: HTMLDivElement = document.createElement("div");
		divCard.className = "card";
		divCard.id = `${attachment.attachmentId.id}_divcard`;
		this._attachmentContainer.appendChild(divCard);

		//set item name
		let divCardBody: HTMLDivElement = document.createElement("div");
		divCardBody.className = "card-text text-left text-truncate";
		divCard.appendChild(divCardBody);
		divCardBody.title = `${attachment.name}.${attachment.extension}`;
		divCardBody.innerHTML = `${attachment.name}.${attachment.extension}`;

		let fileIcon = this.getFileIcon(attachment);
		let divFileIcon: HTMLDivElement = document.createElement("div");
		divFileIcon.classList.add("card-img-top", "fileIconDiv");
		divFileIcon.innerHTML = `<span class="fileIcon"><i title="${attachment.name}.${attachment.extension}" class='${fileIcon}'></i></span>`;
		divCard.appendChild(divFileIcon);

		// let divFileDetails: HTMLDivElement = document.createElement("div");
		// divFileDetails.classList.add("float", "fileDetailsDiv");
		// divFileDetails.innerHTML = `<span class="fileIcon"><i title="${attachment.name}.${attachment.extension}" class='${fileIcon}'></i></span>`;
		// divCard.appendChild(divFileDetails); 
		
		
		//get item image
		// let img: HTMLImageElement = <HTMLImageElement>document.createElement("img");
		// img.className = "card-img-top";
		// divCard.appendChild(img);
		// this.findImage(img, attachment);

		let divButtons: HTMLDivElement = document.createElement("div");
		divButtons.id = `${attachment.attachmentId.id}_divbuttons`;
		divButtons.className = "divButtons";
		divCard.appendChild(divButtons);
		//create download element
		let downloadButton: HTMLButtonElement = document.createElement("button");
		downloadButton.type = "button";
		downloadButton.className = "close downloadButton";
		downloadButton.id = `${attachment.attachmentId.id}_downloadButton`;
		downloadButton.innerHTML = `<span><i title='Download ${attachment.name}.${attachment.extension}' class='fas fa-download'></i></span>`;
		downloadButton.title = `Download ${attachment.name}.${attachment.extension}`;
		divButtons.appendChild(downloadButton);

		let divFileSize: HTMLSpanElement = document.createElement("div");
		divFileSize.className = "fileSize";
		divFileSize.title = "File Size";
		divFileSize.innerHTML = `${this.readableBytes(attachment.size)}`;
		divButtons.appendChild(divFileSize);

		//create delete element
		let deleteButton: HTMLButtonElement = document.createElement("button");
		deleteButton.type = "button";
		deleteButton.className = "close deleteButton";
		deleteButton.id = `${attachment.attachmentId.id}_deleteButton`;
		//deleteButton.innerHTML = "<span><i class='fas fa-trash-alt'></i></span>";
		deleteButton.innerHTML = `<span><i title='Delete ${attachment.name}.${attachment.extension}' class='far fa-trash-alt'></i></span>`;
		deleteButton.title = `Delete ${attachment.name}.${attachment.extension}`;
		divButtons.appendChild(deleteButton);

		
		let attachmentRef = new AttachmentRef(
			attachment.attachmentId.id.toString(),
			attachment.attachmentId.typeName);

		//add event listeners 
		divCardBody.addEventListener("click", this.onClickAttachment.bind(this, attachmentRef, undefined));
		divFileIcon.addEventListener("click", this.onClickAttachment.bind(this, attachmentRef, undefined));
		deleteButton.addEventListener("click", this.onClickDelete.bind(this, divCard, attachmentRef, this._attachmentSource));
		downloadButton.addEventListener("click", this.onClickAttachment.bind(this, attachmentRef, { openMode: 2 }));
	}

	private onClickDelete(divCard: HTMLDivElement, attachment: AttachmentRef) {
		//show confirm or cancel action
		let confirmDelete: HTMLDivElement = document.createElement("div");
		confirmDelete.className = "confirmDelete";
		divCard.appendChild(confirmDelete);
		//add html button group
		confirmDelete.innerHTML = `<div class='btn-group btn-group-sm' role='group'><button alt='Delete' id='deleteButton' type='button' class='btn btn-secondary'>Delete</button><button alt='Cancel' id='cancelButton' type='button' class='btn btn-secondary'>Cancel</button></div>`;

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

	private readableBytes(bytes: number): string {
		let i: number = Math.floor(Math.log(bytes) / Math.log(1024)),
		sizes = ['B', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB'];
	
		return `${(bytes / Math.pow(1024, i)).toFixed(1)} ${sizes[i]}`;
	}

	private removeBSCard(id: string) {
		let fileElement = document.getElementById(`${id}_divcard`);
		if (fileElement) {
			fileElement.remove();
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
			this.removeBSCard(result.id);
		})
	}

	private onClickCancelDelete(confirmDelete: HTMLDivElement) {
		//clear the buttons
		confirmDelete.innerHTML = "";
	}

	private getFileIcon(item: Attachment): string{
		let result = "far fa-file-alt";
		switch(item.extension.toLocaleLowerCase()){
			case "pdf":
				result = "far fa-file-pdf";
				break;
			case "doc":
			case "docx":
				result = "far fa-file-word";
				break;
			case "ppt":
			case "pptx":
				result = "far fa-file-powerpoint";
				break;
			case "xls":
			case "xlsx":
				result = "far fa-file-excel";
				break;
			case "css":
			case "html":
				result = "far fa-file-code";
				break;
			default:
				result = "far fa-file-alt"
				break;
		}

		return result;
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
				let item = new Attachment(new EntityReference("annotation", ent.annotationid.toString()), ent.filename.split('.')[0], ent.filename.split('.')[1].toLowerCase(), false, ent.filesize);

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
				file.fileContent = type == "annotation" ? result.documentbody : result.body;
				file.fileName = result.filename;
				file.fileSize = result.filesize;
				file.mimeType = result.mimetype;
				return file;
			}

		);
	}

	private onClickAttachment(attachment: AttachmentRef, fileOptions?: OpenFileOptions): void {
		//get the attachment id
		let id = attachment.id;
		let type = attachment.type;

		this.getAttachment(id, type).then((r: any) => {
			this._context.navigation.openFile(r, fileOptions);
		}
		);
	}

	private async initAttachmentsAndDropHandler(context: ComponentFramework.Context<IInputs>) {
		if ((<any>context).page.entityId != null) {
			this._attachmentsNotAvialbleElement.classList.add("hidden");
			this._reference = new EntityReference(
				(<any>context).page.entityTypeName,
				(<any>context).page.entityId			
			)
			await this.getAttachments(this._reference)
			this._dropHandler = new DropHandler(this._apiClient, this._progressElement, this._progressBar, this._attachmentSource, this._refreshFileNameAfterUpload);
			this._dropHandler.HandleDrop(this._dropElement, (<any>context).page.entityId, (<any>context).page.entityTypeName);
		}
		else{
			this._attachmentsNotAvialbleElement.classList.remove("hidden");
		}	
	}

	public async updateView(context: ComponentFramework.Context<IInputs>) {
		if (!this._dropHandler)
		{				
			await this.initAttachmentsAndDropHandler(context);
		}
	}

	public getOutputs(): IOutputs {
		return {};
	}

	public destroy(): void {
		// Add code to cleanup control if necessary

	}

}