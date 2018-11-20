import { dateAdd } from "@pnp/common";
import { List } from "@pnp/sp";
import { ITemplateItem } from "./ITemplateItem";
import { TemplateItemParser } from "./TemplateItemParser";

export class TemplatesList extends List {
  public static Title: string = 'MD Templates';
  public static ListeItemType: string = 'SP.Data.MD_x0020_TemplatesListItem';

  public getTemplates(): Promise<ITemplateItem[]> {

    return this.items.select('Title', 'Template', 'Id').usingCaching({
      expiration: dateAdd(new Date(), "minute", 10),
      key: 'md_templates'
    }).get(new TemplateItemParser());
  }
}
