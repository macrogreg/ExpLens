# ExpLens Excel-AddIn

## Licenses

### Intent

ExpLens enables you to work with financial data freely and without restrictions, privately or commercially. You can use ExpLense free of charge for your family, your organization, or _as a tool_ in services you offer.

However, to protect our investment in building and maintaining ExpLense, we restrict using it for _competitive_ or _alternative_ offerings to our core mission:  
  » Tools for financial data analysis  
  » Synchronization of financial data

**Interested in collaborating on a similar offering?  
Reach out! We'll work out reasonable terms.**

### Summary

The above the intent of our licensing strategy.
See [LICENSE.md](./LICENSE.md) for complete terms.

- **Generic libraries** (explicitly listed): MIT License
    - ✓ use freely, including commercially.

- **Everything else**: PolyForm Shield 1.0.0 License
    - ✓ Personal and internal use: _allowed_.
    - ✓ Reading, modifying, and sharing code: _allowed_.
    - ✗ Building competing commercial products/services: _not allowed_.
    - ✗ Selling or hosting as a service: _not allowed_.

_Disclaimer: The authors and contributors accept no responsibility, direct or indirect, for any consequences from
using this project._

## Developer guide

(If some info was missing and you needed to figure it out, please create a pull request to add it here.)

#### Install

Clone the repo. Then:

`npm install` or `npm ci`

#### Build and start locally

`npm run dev --qPublicBasePath=app/current/excel`  
(this invokes `quasar dev` and passes a parameter to set the URL base the way Excel expects it)

Quasar will build the app and host it under `http://localhost:9000/app/current/excel/`.  
Note: if port 9000 is taken, Quasar might silently choose another port. However, Excel will look for ExpLens on
port 9000. If you need to use another port, make sure to modify the manifest accordingly
(`ExpLens.Excel-AddIn.Manifest.xml`).

##### Build for production

`quasar build` or `npm run build --qPublicBasePath=url/base/path`

#### Activate the Excel Add-In

The _release_ version of ExpLens works on Excel Web and Excel Desktop (Windows & Mac).  
The instructions here are for the _dev_ version (Windows desktop).

##### Option 1: "Admin Style" (easer to uninstall)

1. Share the root folder of your cloned repo as a local network drive.  
   E.g., if you cloned into `c:\Code\ExpLens\` and your machine is called `DEV-PC`, then your share might be
   called `\\DEV-PC\ExpLens`.

2. In Excel, add the network share as a trusted Add-In Catalog:
    - Go to: _File_ > _Options_ > _Trust Center_ > _Trust Center Settings_ > _Trusted Add-In Catalogs_.
    - In _Catalog Url_, enter the address of your file share
      (e.g., `\\DEV-PC\ExpLens`).
    - Click _Add catalog_, select _Show in menu_, OK all dialogs, and restart Excel (make sure to close all windows).

3. Activate the Add-In:  
   _Home_ > _Add-ins_ > _More Add-ins_. At the top you will have a _Shared Folder_ option. When you select it, you
   will see the ExpLens add-in in the list. Install it.

The Add-In will start and attempt to load from
`http://localhost:9000/app/current/excel/`

(See section 'Build and start locally' above.)

###### Uninstalling:

Return to _Trusted Add-In Catalogs_, remove the shared repo directory, conform, and restart Excel (all windows).

##### Option 2: "Dev Style" (easer to get started)

1. Clone the repo and install packages.  
   (see above)

2. Build and start the dev server.  
   (see above)

3. Install the manifest using the Office AddIn tool:  
   `npm run office-start`  
   This will put some settings into the registry and open Excel with the Add-In loaded.

The Add-In will start and attempt to load from
`http://localhost:9000/app/current/excel/`

(See section 'Build and start locally' above.)

###### Uninstalling:

If you thought this was easier than the "Admin Style", then you have not yet tried to uninstall the
Add-In (which is a normal par of the dev process).
If you time it well, then this will do it:  
`npm run office-stop`

But some of the time the state required for the tool to correctly know how to uninstall is lost, and you have to
go to the registry. The add-ins are listed under:
`Computer\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer`.
Find the one you want to remove, and delete all associated entries (there are several).
