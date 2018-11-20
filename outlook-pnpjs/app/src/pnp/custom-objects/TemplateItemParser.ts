import { ODataParserBase } from "@pnp/odata";
import { ITemplateItem } from "./ITemplateItem";

export class TemplateItemParser extends ODataParserBase<ITemplateItem[]> {

  public parse(response: Response): Promise<ITemplateItem[]> {

    return super.parse(response)
      .then((data: any[]) => {
        return data.map(d => {
          return {
            displayName: d.Title,
            template: d.Template
          };
        });
      });
  }
}
