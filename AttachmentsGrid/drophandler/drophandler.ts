import { fromEvent } from 'rxjs';

export class DropHandler {

    private _entity: ComponentFramework.WebApi.Entity;

    constructor(private webApi: ComponentFramework.WebApi) {

    }

    public HandleDrop(target: HTMLElement) {
        fromEvent(target, 'drop')
            .subscribe(async (ev: any) => {
                ev.preventDefault();
                if (ev.dataTransfer.items) {
                    await this._HandleFiles(ev.dataTransfer.files, "1", "contact");
                }
            });

        fromEvent(target, 'dragover')
            .subscribe(ev => ev.preventDefault());
    }

    private async _HandleFiles(list: FileList, entityId: string, entityLogicalName: string) {

        console.log(list);
        console.log(`${list.length} files detected`)

        for (var i = 0; i < list.length; i++) {
            console.log(`File No: ${i} processing of ${list.length}`);
            var file = list[i];

            var encodedData = await this._EncodeFile(file);

            var attachment = {
                "subject": file.name,
                "filename": file.name,
                "documentbody": encodedData
            }

            console.log(`Processing File: ${attachment.filename}`);

            console.log("Neeeext!");
        }
    }

    private _EncodeFile(file: File): Promise<string> {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (f) => resolve((<string>reader.result).split(',')[1]);
            reader.onerror = error => reject(error);
            reader.readAsDataURL(file);
        });
    }

}