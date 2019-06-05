import { IInputs, IOutputs } from "./generated/ManifestTypes";
import 'bootstrap';
import { DropHandler } from "./drophandler/drophandler";
import { Attachment } from "./Attachment";
import { EntityReference } from "./EntityReference";

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

	private _context: ComponentFramework.Context<IInputs>;
	private _container: HTMLDivElement;

	// Define Input Elements
	public _dropElement: HTMLDivElement;

	public _progressElement: HTMLDivElement;
	public _progressBar: HTMLDivElement;

	private _dropHandler: DropHandler;

	private _apiClient: ComponentFramework.WebApi;

	private _attachments: Attachment[];

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
		this._apiClient = context.webAPI;

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

		this._progressElement.append(this._progressBar);

		this._container.append(this._progressElement, this._dropElement);

		// Bind to parent container
		container.append(this._container);

		//get attachements from notes
		let reference: EntityReference = new EntityReference(
			(<any>context).page.entityTypeName,
			(<any>context).page.entityId
		)
		if ((<any>context).page.entityId != null) {
			this._attachments = await this.getAttachments(reference)
			this.createBSCards();
		}

		// TODO: Remove this:
		// this._attachments = await this.getAttachments(new EntityReference('jojhn', '1'));
		// this.createBSCards(this._attachments);

		this._dropHandler = new DropHandler(this._apiClient, this._progressElement, this._progressBar, this._attachments);
		this._dropHandler.HandleDrop(this._dropElement, (<any>context).page.entityId, (<any>context).page.entityTypeName);
	}

	private createBSCards() {
		//create the bootstrap cards
		if (this._attachments.length > 0) {
			//create containing card
			let divControl: HTMLDivElement = document.createElement("div");
			divControl.className = "card-columns";
			this._dropElement.appendChild(divControl);

			this._attachments.forEach(item => {
				//create item card
				let divCard: HTMLDivElement = document.createElement("div");
				divCard.className = "card";
				divControl.appendChild(divCard);

				//get item image
				let img: HTMLImageElement = <HTMLImageElement>document.createElement("img");
				img.className = "card-img-top";
				divCard.appendChild(img);
				this.findImage(img, item);

				//set item name
				let divCardBody: HTMLDivElement = document.createElement("div");
				divCardBody.className = "card-text text-center text-truncate";
				divCard.appendChild(divCardBody);

				divCardBody.innerHTML = `${item.name}.${item.extension}`;
				let attachmentRef = new AttachmentRef(
					item.attachmentId.id.toString(),
					item.attachmentId.typeName);

				//add a view file listener
				divCard.addEventListener("click", this.onClickAttachment.bind(this, attachmentRef));

			})
	}
}

	private findImage(img: HTMLImageElement, item: Attachment) {
		//find the image
		this._context.resources.getResource(`${item.extension}.png`,
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
				let item = new Attachment(new EntityReference("annotation", ent["annotationid"].toString()), ent["filename"].split('.')[0], ent["filename"].split('.')[1].toLowerCase());
				items.push(item);
			}
			return items;
		}
		catch (error) {
			console.log(error.message);
			let items_1: Attachment[] = [];
			return items_1;
		}

		//return [new Attachment(new EntityReference('annotation', '1'), 'Document', 'png'), new Attachment(new EntityReference('annotation', '2'), 'Document2', 'png')]

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

	}

	public getOutputs(): IOutputs {
		return {};
	}

	public destroy(): void {
		// Add code to cleanup control if necessary

	}

}