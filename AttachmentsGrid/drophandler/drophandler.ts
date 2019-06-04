import { fromEvent } from 'rxjs';
import { Attachment } from '../Attachment';
import { EntityReference } from '../EntityReference';

export class DropHandler {

    private _entity: ComponentFramework.WebApi.Entity;

    constructor(private webApi: ComponentFramework.WebApi, private progressElement: HTMLDivElement, private progressBar: HTMLDivElement, private attachments: Attachment[]) {

    }

    public HandleDrop(target: HTMLElement, entityId: string, entityLogicalName: string) {
        fromEvent(target, 'drop')
            .subscribe(async (ev: any) => {
                ev.preventDefault();
                if (ev.dataTransfer.items) {
                    ev.target.classList.remove('drop-zone-hover');
                    this.progressElement.style.visibility = "visible";
                    await this.HandleFiles(ev.dataTransfer.files, entityId, entityLogicalName);
                    
                    // This is to show the progress bar going to 100% before hiding, for better UX.
                    await this.sleep(750);
                    this.progressElement.style.visibility = "hidden";
                    this.progressBar.style.width = "0%";
                }
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

    private async HandleFiles(list: FileList, entityId: string, entityLogicalName: string) {

        for (var i = 0; i < list.length; i++) {

            var file = list[i];

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

            this.attachments.push(new Attachment(new EntityReference('annotation', response.id), file.name, this.GetFileExtension(file.name)));

            console.log(`Attachment: ${file.name} processed, percentage complete: ${this.GetProgressPercentage(i + 1, list.length)}`);

            this.progressBar.style.width = this.GetProgressPercentage(i + 1, list.length);

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

    private GetProgressPercentage(complete: number, todo: number) {
        return `${Math.round((complete / todo) * 100)}%`;
    }

    private async sleep(msec: number) {
        return new Promise(resolve => setTimeout(resolve, msec));
    }
}