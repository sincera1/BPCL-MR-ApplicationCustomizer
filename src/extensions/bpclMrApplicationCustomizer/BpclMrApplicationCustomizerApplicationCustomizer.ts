import { Log } from '@microsoft/sp-core-library';
import styles from './BpclMrApplicationCustomizer.module.scss';
import Logo from './assets/Bpcl-logo.png';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

const LOG_SOURCE =
  'BpclMrApplicationCustomizer';

const translations: any = {

  English: {

    home: "HOME",
    bpmai: "BPMAI",
    myPortal: "MY PORTAL",
    forms: "FORMS",
    refineryVision: "REFINERY VISION",
    iconnect: "I-CONNECT",
    erdmp: "ERDMP",
    zohoRemote: "ZOHO REMOTE",
    afpm: "AFPM"

  },

  Hindi: {

    home: "होम",
    bpmai: "बीपीएमएआई",
    myPortal: "माय पोर्टल",
    forms: "फॉर्म्स",
    refineryVision: "रिफाइनरी विज़न",
    iconnect: "आई-कनेक्ट",
    erdmp: "ईआरडीएमपी",
    zohoRemote: "ज़ोहो रिमोट",
    afpm: "एएफपीएम"

  },

  Marathi: {

    home: "होम",
    bpmai: "बीपीएमएआई",
    myPortal: "माय पोर्टल",
    forms: "फॉर्म्स",
    refineryVision: "रिफायनरी व्हिजन",
    iconnect: "आय-कनेक्ट",
    erdmp: "ईआरडीएमपी",
    zohoRemote: "झोहो रिमोट",
    afpm: "एएफपीएम"

  }

};

export default class BpclMrApplicationCustomizerApplicationCustomizer extends BaseApplicationCustomizer<{}> {
  

  private _top:PlaceholderContent | undefined;
    

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

  private _hideAppBar():
    void {

    const style =
      document.createElement(
        'style'
      );

    style.innerHTML = `

      #spSiteHeader {
        display: none !important;
      }

      #sp-appBar {
        display: none !important;
      }

      .ms-HorizontalNav {
        margin-left: 0 !important;
      }

      #spCommandBar {
        display: none !important;
      }

      #vpc_Page\\.SiteFooter\\.internal\\.03025612-a400-4804-a78e-e1493200a43b {
        display: none !important;
      }

      #CommentsWrapper {
        display: none !important;
      }

    `;

    document.head.appendChild(
      style
    );

  }

  private async _renderTop():Promise<void> {
    

     if (document.getElementById('bpcl-custom-navbar')) {
  return;
}


    this._top =this.context.placeholderProvider .tryCreateContent(PlaceholderName.Top );
      

    if (
      !this._top ||
      !this._top.domElement
    ) {
      return;
    }

    const selectedLanguage =
      localStorage.getItem(
        "MRLanguage"
      ) || "English";

    const t =
      translations[
      selectedLanguage
      ];

    this._top.domElement.innerHTML = `

<div id="bpcl-custom-navbar" class="${styles.mainNavbar}">

<div class="${styles.navbarContainer}">

<div class="${styles.logoSection}">

<img src="${Logo}" alt="BPCL Logo" class="${styles.logoImg}" />

<span class="${styles.logoText}">MR Connect</span>

</div> 

<button
      id="bpclMenuBtn"
      class="${styles.menuBtn}"
      aria-label="Toggle Menu">
      ☰
    </button>

<div id="bpclMenu" class="${styles.navMenu}">

<div class="${styles.navLinks}">

<a href="#" class="${styles.navLink}">${t.home}</a>

<a href="#" class="${styles.navLink}">${t.bpmai}</a>

<a href="#" class="${styles.navLink}">${t.myPortal}</a>

<a href="#" class="${styles.navLink}">${t.forms}</a>

<a href="#" class="${styles.navLink}">${t.refineryVision}</a>

<a href="#" class="${styles.navLink}">${t.iconnect}</a>

<a href="#" class="${styles.navLink}">${t.erdmp}</a>

<a href="#" class="${styles.navLink}">${t.zohoRemote}</a>

<a href="#" class="${styles.navLink}">${t.afpm}</a>

</div>

<div class="${styles.languageWrapper}">

<select id="mrLanguageSelector" class="${styles.languageSelect}">

<option
value="English"
${selectedLanguage === "English"
        ? "selected"
        : ""}
>
English
</option>

<option
value="Hindi"
${selectedLanguage === "Hindi"
        ? "selected"
        : ""}
>
Hindi
</option>

<option
value="Marathi"
${selectedLanguage === "Marathi"
        ? "selected"
        : ""}
>
Marathi
</option>

</select>

</div>

</div>

</div>

</div>

`;

setTimeout(() => {
 
  const menuBtn =
    document.getElementById('bpclMenuBtn');
 
  const menu =
    document.getElementById('bpclMenu');
 
  if (menuBtn && menu) {
 
    menuBtn.addEventListener('click', () => {
 
      const isOpen =
        menu.classList.contains('bpcl-mobile-open');
 
      if (isOpen) {
 
        menu.classList.remove('bpcl-mobile-open');
        menuBtn.textContent = '☰';
 
      } else {
 
        menu.classList.add('bpcl-mobile-open');
        menuBtn.textContent = '✕';
 
      }
 
    });
 
  }
 
}, 300);

    const dropdown =
      document.getElementById(
        "mrLanguageSelector"
      ) as HTMLSelectElement;

    if (dropdown) {

      dropdown.addEventListener(
        "change",
        async (e: any) => {

          localStorage.setItem(
            "MRLanguage",
            e.target.value
          );

          if (this._top) {

            this._top.dispose();

            this._top = undefined;

          }

          await this._renderTop();

          window.dispatchEvent(
            new CustomEvent(
              "languageChanged"
            )
          );

        }
      );

    }

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

