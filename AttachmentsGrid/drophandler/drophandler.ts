import { fromEvent } from 'rxjs';

export class DropHandler {

    private _entity: ComponentFramework.WebApi.Entity;

    constructor (private webApi: ComponentFramework.WebApi) {

    }

    public HandleDrop(target: HTMLElement) {
        fromEvent(target, 'drop')
        .subscribe(async (ev: any) => {
            ev.preventDefault();
            if (ev.dataTransfer.items) {
                await this._HandleFiles(ev.dataTransfer.items, "1", "contact");
            }
        });

        fromEvent(target, 'dragover')
        .subscribe(ev => ev.preventDefault());
    }

    private async _HandleFiles(list: DataTransferItemList, entityId: string, entityLogicalName: string) {
        for (var i = 0; i < list.length; i++) {
            if (list[i].kind === 'file') {
                var file = <File>list[i].getAsFile();

                var encodedData = await this._EncodeFile(file);

                var attachment = {
                    "subject": file.name,
                    "filename": file.name,
                    "documentbody": encodedData
                }

                console.log(attachment);

                //this.webApi.createRecord("annotation",thing);
            }
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