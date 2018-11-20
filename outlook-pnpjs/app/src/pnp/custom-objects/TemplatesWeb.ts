import { sp, Web } from "@pnp/sp";
import { DefaultFiles } from "../DefaultFiles";
import { TemplatesList } from "./TemplatesList";

/* tslint:disable:no-floating-promises */
export class TemplatesWeb extends Web {

  public TemplatesList(): Promise<TemplatesList> {
    return sp.web.lists.ensure(TemplatesList.Title, 'MD Templates library', 100)
      .then(ensureResult => {
        if (ensureResult.created) {
          return ensureResult.list.fields.addMultilineText('Template', 6, false)
            .then(() => {
              return ensureResult.list.defaultView.fields.add('Template');
            })
            .then(() => {
              const batch = sp.createBatch();

              ensureResult.list.items.inBatch(batch).add({
                'Title': 'Hello World',
                'Template': DefaultFiles.HelloWorld
              }, TemplatesList.ListeItemType);

              ensureResult.list.items.inBatch(batch).add({
                'Title': 'Code Sample',
                'Template': DefaultFiles.CodeSample
              }, TemplatesList.ListeItemType);

              ensureResult.list.items.inBatch(batch).add({
                'Title': 'Styled Readme',
                'Template': DefaultFiles.StyledReadme
              }, TemplatesList.ListeItemType);

              ensureResult.list.items.inBatch(batch).add({
                'Title': 'Images',
                'Template': DefaultFiles.Images
              }, TemplatesList.ListeItemType);

              ensureResult.list.items.inBatch(batch).add({
                'Title': 'Emojies',
                'Template': DefaultFiles.Emojies
              }, TemplatesList.ListeItemType);

              return batch.execute();
            })
            .then(() => {
              return ensureResult.list.as(TemplatesList);
            });
        } else {
          return ensureResult.list.as(TemplatesList);
        }
      });
  }
}
