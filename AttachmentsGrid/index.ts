import { IInputs, IOutputs } from "./generated/ManifestTypes";
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
type DataSet = ComponentFramework.PropertyTypes.DataSet;
import 'bootstrap';
import { DropHandler } from "./drophandler/drophandler";

class EntityReference {
	id: string;
	typeName: string;
	constructor(typeName: string, id: string) {
		this.id = id;
		this.typeName = typeName;
	}
}

class Attachment {
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

	private _refreshData: EventListenerOrEventListenerObject;

	private _dropHandler: DropHandler;

	private _apiClient: ComponentFramework.WebApi;

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
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {

		// Add control initialization code
		this._context = context;
		this._container = container;
		this._apiClient = context.webAPI;

		// Layout Elements
		this._dropElement = document.createElement("div");
		this._dropElement.classList.add("drop-zone");

		this._dropHandler = new DropHandler(this._apiClient);
		this._dropHandler.HandleDrop(this._dropElement);

		//get attachements from notes
		let reference: EntityReference = new EntityReference(
			(<any>context).page.entityTypeName,
			(<any>context).page.entityId
		)
		if ((<any>context).page.entityId != null) {
			this.getAttachments(reference).then(r => this.createBSCards(r));
		}
	}

	private createBSCards(items: Attachment[]) {
		//create the bootstrap cards
		if (items.length > 0) {
			//create containing card
			let divControl: HTMLDivElement = document.createElement("div");
			divControl.className = "card-columns";
			this._container.appendChild(divControl);

			for (let i = 0; i < items.length; i++) {
				//create item card
				let divCard: HTMLDivElement = document.createElement("div");
				divCard.className = "card";
				divControl.appendChild(divCard);

				//get item image
				let img: HTMLImageElement = <HTMLImageElement>document.createElement("img");
				img.className = "card-img-top";
				divCard.appendChild(img);
				this.findImage(img, items[i]);

				//set item name
				let divCardBody: HTMLDivElement = document.createElement("div");
				divCardBody.className = "card-text text-center";
				divCard.appendChild(divCardBody);

				divCardBody.innerHTML = items[i].name + "." + items[i].extension;
				let attachmentRef = new AttachmentRef(
					items[i].attachmentId.id.toString(),
					items[i].attachmentId.typeName);

				//add a view file listener
				divCard.addEventListener("click", this.onClickAttachment.bind(this, attachmentRef));

			}
		}
	}

	private findImage(img: HTMLImageElement, item: Attachment) {
		//find the image
		let res = this._context.resources.getResource(item.extension + ".png",
			this.setImage.bind(this, img, "png"),
			this.showError.bind(this)
		);
	}

	private setImage(element: HTMLImageElement, fileType: string, fileContent: string): void {
		//set the image to the img element
		fileType = fileType.toLowerCase();
		let imageUrl: string = "data:image/" + fileType + ";base64, " + fileContent;
		element.src = imageUrl;
	}

	private showError(): void {
		console.log('error while downloading .png');
	}

	private getAttachments(reference: EntityReference): Promise<Attachment[]> {
		//webapi query to find any attachments
		let query = "?$select=annotationid,filename,filesize,createdon,mimetype&$filter=filename ne null and _objectid_value eq " + reference.id + " and objecttypecode eq '" + reference.typeName + "' &$orderby=createdon desc";
		return this._context.webAPI.retrieveMultipleRecords("annotation", query).then(
			function success(result) {
				let items: Attachment[] = [];
				for (let i = 0; i < result.entities.length; i++) {
					let ent = result.entities[i];
					let item = new Attachment(
						new EntityReference("annotation", ent["annotationid"].toString()),
						ent["filename"].split('.')[0],
						ent["filename"].split('.')[1].toLowerCase());
					items.push(item);

				}
				return items;
			}
			, function (error) {
				console.log(error.message);
				let items: Attachment[] = [];
				return items;
			}
		);

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