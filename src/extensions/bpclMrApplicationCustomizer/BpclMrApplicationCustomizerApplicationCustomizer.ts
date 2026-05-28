import { Log } from '@microsoft/sp-core-library';

import styles from './BpclMrApplicationCustomizer.module.scss';
import Logo from './assets/Bpcl-logo.png';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

const LOG_SOURCE = 'BpclMrApplicationCustomizer';

export default class BpclMrApplicationCustomizerApplicationCustomizer
  extends BaseApplicationCustomizer<{}> {

  private _top: PlaceholderContent | undefined;

  public async onInit(): Promise<void> {

    Log.info(
      LOG_SOURCE,
      'Initialized'
    );

    this._hideAppBar();

    await this._renderTop();

    this._renderBottom();

    this.context.application.navigatedEvent.add(
      this,
      () => {
        this._renderBottom();
      }
    );

    return Promise.resolve();
  }

  private _hideAppBar(): void {

    const style = document.createElement(
      'style'
    );

    style.innerHTML = `
       #spSiteHeader { display: none !important; }
      #sp-appBar { display: none !important; }
      .ms-HorizontalNav { margin-left: 0 !important; }
      #spCommandBar { display: none !important; }
      #vpc_Page\\.SiteFooter\\.internal\\.03025612-a400-4804-a78e-e1493200a43b { display: none !important; }
      #CommentsWrapper { display: none !important; }
      
    `;

    document.head.appendChild(style);
  }

  private async _renderTop():
    Promise<void> {

    if (this._top)
      return;

    this._top =
      this.context.placeholderProvider
        .tryCreateContent(
          PlaceholderName.Top
        );

    if (
      !this._top ||
      !this._top.domElement
    ) {
      return;
    }

    this._top.domElement.innerHTML = `

<div class="${styles.mainNavbar}">

<div class="container-fluid">

<div
style="
display:flex;
align-items:center;
justify-content:space-between;
flex-wrap:wrap;
"
>

<div
style="
display:flex;
align-items:center;
"
>

<img
src="${Logo}"
alt="BPCL Logo"
class="${styles.logoImg}"
>

<span class="${styles.logoText}">
MR Connect
</span>

</div>

<div
style="
display:flex;
align-items:center;
gap:20px;
flex-wrap:wrap;
"
>

<a href="#"
class="${styles.navLink}">
HOME
</a>

<a href="#"
class="${styles.navLink}">
BPMAI
</a>

<a href="#"
class="${styles.navLink}">
MY PORTAL
</a>

<a href="#"
class="${styles.navLink}">
FORMS
</a>

<a href="#"
class="${styles.navLink}">
REFINERY VISION
</a>

<a href="#"
class="${styles.navLink}">
I-CONNECT
</a>

<a href="#"
class="${styles.navLink}">
ERDMP
</a>

<a href="#"
class="${styles.navLink}">
ZOHO REMOTE
</a>

<a href="#"
class="${styles.navLink}">
AFPM
</a>

</div>

<div class="${styles.languageWrapper}">

<select
class="${styles.languageSelect}"
>

<option>
English
</option>

<option>
Tamil
</option>

<option>
Hindi
</option>

</select>

</div>

</div>

</div>

</div>

`;

  }

  private _renderBottom():
    void {

    if (
      document.getElementById(
        'bpcl-footer'
      )
    ) {
      return;
    }

    const pageCanvas =
      document.querySelector(
        '#spPageCanvasContent'
      );

    if (!pageCanvas) {
      return;
    }

    const year =
      new Date()
        .getFullYear();

    const footer =
      document.createElement(
        'div'
      );

    footer.id =
      'bpcl-footer';

    footer.innerHTML = `

<div class="${styles.footerSection}">

<div class="${styles.footerContent}">

<div class="${styles.footerLeft}">
© ${year}
Bharat Petroleum Corporation Limited
</div>

<div class="${styles.footerRight}">
<a href="#">Feedback</a>
<a href="#">Help</a>
</div>

</div>

</div>

`;

    pageCanvas.appendChild(
      footer
    );

  }

}