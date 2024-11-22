const { chromium } = require("playwright")

;(async () => {
    const userDataDir = "C:/Users/lolre/AppData/Local/Google/Chrome/User Data"
    const url =
        "https://soneparprod.service-now.com/now/nav/ui/classic/params/target/%24pa_dashboard.do%3Fsysparm_dashboard%3D5fb6e1a2c3386d94c354254ce00131a1%26sysparm_tab%3D11c6e5a2c3386d94c354254ce001316e%26sysparm_cancelable%3Dtrue%26sysparm_editable%3Dundefined%26sysparm_active_panel%3Dfalse"

    const browser = await chromium.launchPersistentContext(userDataDir, {
        headless: false,
        executablePath: "C:/Program Files/Google/Chrome/Application/chrome.exe",
    })
    const page = await browser.newPage()
    await page.goto(url, { waitUntil: "domcontentloaded" })
    await page.getByPlaceholder("Pesquisar").click()
    await page.getByPlaceholder("Pesquisar Global").fill("INC0668760")
    await page.getByPlaceholder("Pesquisar Global").press("Enter")
    await page.getByLabel("CORRESPONDÊNCIA EXATA (1 de 1)").getByRole("button", { name: "NCI ATRELADA A UMA" }).click()
    await page.locator('iframe[name="gsft_main"]').contentFrame().getByLabel("menu de ações adicionais").click()
    await page.locator('iframe[name="gsft_main"]').contentFrame().getByRole("menuitem", { name: "Histórico " }).click()
    await page.locator('iframe[name="gsft_main"]').contentFrame().getByRole("menuitem", { name: "Calendário" }).click()
    await page.waitForSelector('iframe[name="gsft_main"]')

    const iframe = await page.frame({ name: "gsft_main" })
    if (!iframe) {
        console.error("Iframe not found!")
        return
    }

    await iframe.waitForSelector('[id="img\\.historylist"]', { state: "visible" })

    const historyListImage = await iframe.locator('[id="img\\.historylist"]')
    await historyListImage.click()

    await page.pause()
})()
