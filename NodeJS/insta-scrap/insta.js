const puppeteer = require('puppeteer');
const fs = require('fs');

(async () => {
   const browser = await puppeteer.launch({
      headless: false, slowMo: 20
   });
   const page = await browser.newPage()
   await page.setViewport({width: 1100, height: 790, deviceScaleFactor:1, isMobile:false})

   const URL = 'https://www.instagram.com/p/CChMVvQgYKK/'
   await page.goto(URL)
   if(page.url() != URL) {await doLogin()}

   page.waitForNavigation()


///--    Carregar todos os Comentários    --///
   var i = 0
   await loadMore('.NUiEW > button:nth-child(1)')

///--    Criar array com os comentários      --///
   const arrobas = await getComments('.MOdxS span a')

///--    Contar @ arrobas     --///
   const count = arrobas.reduce((arroba, freq) => arroba.set(freq, (arroba.get(freq) || 0) +1), new Map())

///--    Ordenar qntd. @ do Maior para o Menor     --///
   const sorted = new Map([...count].sort((value1, value2) => value2[1] - value1[1]))

   fs.writeFile('comments.json', JSON.stringify(Object.fromEntries(sorted), null, 2), err => {
      if(err) throw new Error('Something went wrong..')
      console.log(' ~ Generated "comments.json" file with all Profiles and their citation frequency ~' )
   })



///--     Funções     --///
   async function doLogin() {
      await page.waitForSelector('input[name="username"]')

      await page.type('input[name="username"]', '<YOUR USERNAME>')
      await page.type('input[name="password"]', '<YOUR PASSWORD>')
      await page.click('button[type="submit"]')

      const manyLogins = await page.$('#slfErrorAlert')
      if (manyLogins) {
         console.info(' ~ Login Suspended: You have logged too many times on this account - wait 5m and try again ~ ');
         process.exit(0);
      }

      await page.waitForNavigation()

      page.waitForXPath('//button[contains(text(), "Agora não")]').then(() => {
         page.click('.cmbtv > button')
      })
   }

   async function loadMore(selector) {
      await page.waitForSelector('.MOdxS span a')
      const inputLog = await page.$('input[name="username"]')
      if (inputLog) doLogin()

      const moreButton = await page.$(selector)
      if (moreButton) {
         i++
         await moreButton.click()
         console.log('%i | Loading Comments...', i)

         await page.waitForSelector(selector, {timeout: 3000}).catch(() => console.log(' ~ All comments were Loaded ~'))
         await loadMore(selector)
      }
   }

   async function getComments(selector) {
      const arrobas = await page.$$eval(selector, links => links.map(link => link.innerText))
      return arrobas
   }

})()