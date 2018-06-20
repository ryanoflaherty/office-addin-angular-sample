import { Component } from '@angular/core';

declare let Word;

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {

  constructor() { }

  clearDocument() {
    Word.run(function (context) {
      const body = context.document.body;
      body.clear();
      return context.sync();
    });
    // .catch(errorHandler);
  }

  insertText() {
    Word.run(function (context) {
      const docBody = context.document.body;
      docBody.insertParagraph('Lorem ipsum', Word.InsertLocation.end);
      return context.sync();
    });
    // .catch(errorHandler);
  }

}
