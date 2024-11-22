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

    const contentFrame = await page.locator('iframe[name="gsft_main"]').contentFrame()

    if (contentFrame) {
        await contentFrame.getByLabel("menu de ações adicionais").click()
        await contentFrame.getByRole("menuitem", { name: "Histórico " }).click()
        await contentFrame.getByRole("menuitem", { name: "Calendário" }).click()

        const iframe = await page.frame({ name: "gsft_main" })
        if (!iframe) {
            console.error("Iframe not found!")
            return
        }

        const historyListImage = await iframe.locator('[id="img\\.historylist"]')
        await historyListImage.waitFor({ state: "visible", timeout: 10000 })
        await historyListImage.click()

        const historyElements = await contentFrame.locator('[id^="historylist"]')
        const elementsCount = await historyElements.count()
        console.log(`Found ${elementsCount} historylist elements.`)

        if (elementsCount > 0) {
            for (let i = 0; i < elementsCount; i++) {
                const historyElement = historyElements.nth(i)

                await historyElement.scrollIntoViewIfNeeded()
                await historyElement.waitFor({ state: "visible", timeout: 10000 })
                await historyElement.click()
                console.log(`Clicked on historylist element ${i}`)

                const servicoCell = await historyElement.locator(':has-text("Serviço")')
                const servicoExists = await servicoCell.count()

                if (servicoExists > 0) {
                    const servicoText = await servicoCell.first().textContent()

                    console.log(`"Serviço" found with text: ${servicoText}`)

                    const regexCentral = /Central de Atendimento/i
                    const regexFornecedor = /Fornecedor SZ/i

                    const hasCentralAtendimento = regexCentral.test(servicoText)
                    const hasFornecedorSZ = regexFornecedor.test(servicoText)

                    if (hasCentralAtendimento && hasFornecedorSZ) {
                        console.log(`Found: Central de Atendimento and Fornecedor SZ`)
                        const fullText = await historyElement.locator('xpath=preceding-sibling::div[@id="historyEventItem"][1]').textContent()

                        const dateTimeRegex = /^\d{4}-\d{2}-\d{2}~ \d{2}:\d{2}:\d{2}/
                        const dateTime = fullText.match(dateTimeRegex)[0]

                        const excelDateTime = dateTime.replace("~", " ") 

                        console.log("Excel-friendly date and time:", excelDateTime)
                    } else {
                        console.log(`"Serviço" does not contain both "Central de Atendimento" and "Fornecedor SZ"`)
                        if (!hasCentralAtendimento) console.log(`"Central de Atendimento" not found.`)
                        if (!hasFornecedorSZ) console.log(`"Fornecedor SZ" not found.`)
                    }
                } else {
                    console.log(`"Serviço" not found in historylist element ${i}`)
                }
            }
        } else {
            console.log("No historylist elements found after waiting.")
        }
    } else {
        console.error("Content frame not found!")
    }

    await page.pause()
})()
