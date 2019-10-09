import { fromEvent, Subject } from 'rxjs';
import { Attachment } from '../Attachment';
import { EntityReference } from '../EntityReference';

export class DropHandler {

    constructor(private webApi: ComponentFramework.WebApi, private progressElement: HTMLDivElement, private progressBar: HTMLDivElement, private attachmentSource: Subject<Attachment>, private refreshFileNameAfterUpload: boolean) {

    }

    public HandleDrop(target: HTMLElement, entityId: string, entityLogicalName: string) {
        if (entityId) {
            fromEvent(target, 'drop')
                .subscribe(async (ev: any) => {
                    ev.preventDefault();
                    await this.handleEvent(ev, entityId, entityLogicalName);
                });

            fromEvent(target, 'dragover')
                .subscribe(ev => ev.preventDefault());

            fromEvent(target, 'dragenter')
                .subscribe((ev: any) => {
                    var element = <HTMLElement>ev.target;
                    element.classList.add('drop-zone-hover');
                });

            fromEvent(target, 'dragleave')
                .subscribe((ev: any) => {
                    var element = <HTMLElement>ev.target;
                    element.classList.remove('drop-zone-hover');
                });
        }
    }

    private async handleEvent(ev: any, entityId: string, entityLogicalName: string) {
        if (ev.dataTransfer.items) {
            ev.target.classList.remove('drop-zone-hover');
            this.progressElement.style.visibility = "visible";
            await this.handleFiles(ev.dataTransfer.files, entityId, entityLogicalName);
            // This is to show the progress bar going to 100% before hiding, for better UX.
            await this.sleep(750);
            this.progressElement.style.visibility = "hidden";
            this.progressBar.style.width = "0%";
        }
    }

    private async handleFiles(list: FileList, entityId: string, entityLogicalName: string) {
        var errors = "";

        for (var i = 0; i < list.length; i++) {

            var file = list[i];

            if (!file.size) {
                errors += `File ${file.name} could not be uploaded because it has no content.\n\n`;
                continue;
            }

            var encodedData = await this.EncodeFile(file);

            var attachment = {
                "subject": `Attachment: ${file.name}`,
                "filename": file.name,
                "filesize": file.size,
                "mimetype": file.type,
                "objecttypecode": entityLogicalName,
                "documentbody": encodedData,
                [`objectid_${entityLogicalName}@odata.bind`]: `/${this.CollectionNameFromLogicalName(entityLogicalName)}(${entityId})`
            };

            var response = await this.webApi.createRecord('annotation', attachment);

            var fileName = file.name;
            if (this.refreshFileNameAfterUpload) {
                var uploadedFile = await this.webApi.retrieveRecord('annotation', response.id, "?$select=filename");
                fileName = uploadedFile.filename;
                console.log(`Attachment: Original file name ${file.name} retrieved as ${fileName} from annotation.`);
            }

            this.attachmentSource.next(new Attachment(new EntityReference('annotation', response.id), this.TrimFileExtension(fileName), this.GetFileExtension(fileName), false));

            console.log(`Attachment: ${fileName} processed, percentage complete: ${this.GetProgressPercentage(i + 1, list.length)}`);

            this.progressBar.style.width = this.GetProgressPercentage(i + 1, list.length);
        }
        
        if (errors.length) {
            await this.sleep(750);
            alert(errors);
        }
    }

    private EncodeFile(file: File): Promise<string> {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (f) => resolve((<string>reader.result).split(',')[1]);
            reader.onerror = error => reject(error);
            reader.readAsDataURL(file);
        });
    }

    private CollectionNameFromLogicalName(entityLogicalName: string): string {
        if (entityLogicalName[entityLogicalName.length - 1] != 's') {
            return `${entityLogicalName}s`;
        } else {
            return `${entityLogicalName}es`;
        }
    }

    private GetFileExtension(fileName: string): string {
        return <string>fileName.split('.').pop();
    }

    private TrimFileExtension(fileName: string): string {
        return <string>fileName.split('.')[0];
    }

    private GetProgressPercentage(complete: number, todo: number) {
        return `${Math.round((complete / todo) * 100)}%`;
    }

    private async sleep(msec: number) {
        return new Promise(resolve => setTimeout(resolve, msec));
    }
}