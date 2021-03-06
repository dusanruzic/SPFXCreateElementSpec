import { sp, } from "@pnp/sp";
import { ClientsideWebpart } from "@pnp/sp/clientside-pages";
import "@pnp/sp/webs";
import SharePointService from '../../../services/SharePoint/SharePointService';

export declare namespace MyClientSideWebpartPropertyTypes {
  /**
   * Properties for People (component id: 7f718435-ee4d-431c-bdbf-9c4ff326f46e)
   */
  interface People {
      layout: "1" | "2";
      persons?: any[];
      description: string;
  }
}
export class PageModelHelper {

    /*
  public static async  getInfos(pagename: string): Promise<string> {

      var resultData: any = await sp.web.lists.getByTitle("Site Pages")
          .items.getById(15)
          .select("Title")
          .get();

      return await resultData.Title;
  }
  */

  public static async createCustomPage(name: string, pagename: number, pageType: string): Promise<string> {
    
    const page = await sp.web.addClientsidePage("elemspec" + pagename + ".aspx");
    //console.log("pagetype" + pageType);
    page.title = name;
    const partDefs = await sp.web.getClientsideWebParts();
    //console.log("case a");
    const section = page.addSection();
    //console.log("section added");

    const column1 = section.addColumn(12);

    // find the definition we want, here by id
    const partDef = partDefs.filter(c => c.Id.toUpperCase() === "{983d600f-615d-4609-940c-199e54b8f496}".toUpperCase());
    //console.log(partDefs);
    // optionally ensure you found the def
    if (partDef.length < 1) {
        // we didn't find it so we throw an error
        //console.log('ops');
        throw new Error("Could not find the web part");
    }
    // create a ClientWebPart instance from the definition
    const part = ClientsideWebpart.fromComponentDef(partDef[0]);

    part.setProperties<MyClientSideWebpartPropertyTypes.People>({
        layout: "2",
        /*
        persons: [
            {
                "id": "i:0#.f|membership|vukasin@jvspdev.onmicrosoft.com",
                "upn": "vukasin@jvspdev.onmicrosoft.com",
                "role": "",
                "department": "",
                "phone": "",
                "sip": ""
            }
        ],
        */
        description: SharePointService.newListItemId + ''
    });
    // add a text control to the second new column
    column1.addControl(part);

    await page.save();
    //console.log("case saved");


    return await "done";
  }
}
