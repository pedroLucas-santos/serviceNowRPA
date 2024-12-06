const { chromium } = require("playwright")
const xlsx = require("xlsx")
const fs = require("fs")
const path = require("path")

const readExcelFile = (filePath) => {
    const workbook = xlsx.readFile(filePath)
    const sheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[sheetName]

    const data = xlsx.utils.sheet_to_json(sheet)

    const incidents = data.map((row) => row.Número).filter(Boolean)
    return { incidents, workbook, sheet, data }
}

const saveToExcelFile = (workbook, data, newColumns, newFilePath) => {
    data.forEach((row) => {
        newColumns.forEach((column) => {
            row[column] = row[column] || ""
        })
    })

    const updatedSheet = xlsx.utils.json_to_sheet(data)

    const sheetName = workbook.SheetNames[0]
    workbook.Sheets[sheetName] = updatedSheet

    xlsx.writeFile(workbook, newFilePath)
}

const reorderColumns = (data, columnOrder) => {
    return data.map((row) => {
        const orderedRow = {}
        columnOrder.forEach((column) => {
            orderedRow[column] = row[column]
        })
        return orderedRow
    })
}

;(async () => {
    const filePath = "C:/Users/P0850/Downloads/relatorios/suporte_sonepar_att_2024-12-04.xlsx"
    const newFilePath = "C:/Users/P0850/Downloads/relatorios/suporte_sonepar_updated.xlsx"

    let workbook, data
    if (fs.existsSync(newFilePath)) {
        const { workbook: existingWorkbook, data: existingData } = readExcelFile(newFilePath)
        workbook = existingWorkbook
        data = existingData
    } else {
        const { workbook: initialWorkbook, data: initialData } = readExcelFile(filePath)
        workbook = initialWorkbook
        data = initialData
    }

    const { incidents } = readExcelFile(filePath)

    const userDataDir = "C:/Users/P0850/AppData/Local/Google/Chrome/User Data"
    const url =
        "https://soneparprod.service-now.com/now/nav/ui/classic/params/target/%24pa_dashboard.do%3Fsysparm_dashboard%3D5fb6e1a2c3386d94c354254ce00131a1%26sysparm_tab%3D11c6e5a2c3386d94c354254ce001316e%26sysparm_cancelable%3Dtrue%26sysparm_editable%3Dundefined%26sysparm_active_panel%3Dfalse"

    const browser = await chromium.launchPersistentContext(userDataDir,{
        headless: false,
        executablePath: "C:/Program Files/Google/Chrome/Application/chrome.exe",
        args: ["--disable-gpu", "--disable-dev-shm-usage", "--disable-software-rasterizer"],
        viewport: { width: 1280, height: 720 },
    })

    const page = await browser.newPage()
    await page.goto(url)

    await page.waitForLoadState("networkidle")

    const newColumnName = "Caixa SZ"
    const diffColumnName = "Aberto x Caixa SZ"

    let index = 0

    for (const incident of incidents) {
        console.log(`Processing incident: ${incident}`)

        if (data[index][newColumnName] && data[index][diffColumnName]) {
            console.log(`Skipping incident ${incident} because columns are already filled.`)
            index++
            continue
        }

        await page.getByLabel("Pesquisa global", { exact: true }).locator("span").first().click()
        await page.getByPlaceholder("Pesquisar").fill(incident)
        await page.getByPlaceholder("Pesquisar").press("Enter")

        const contentFrame = await page.locator('iframe[name="gsft_main"]').contentFrame()

        if (contentFrame) {
            await contentFrame.getByLabel("menu de ações adicionais").click()

            let ariaExpanded = await contentFrame.getByLabel("menu de ações adicionais").getAttribute("aria-expanded")
            let retryClick = false

            while (!retryClick) {
                if (ariaExpanded !== "true") {
                    await contentFrame.getByLabel("menu de ações adicionais").click()
                    ariaExpanded = await contentFrame.getByLabel("menu de ações adicionais").getAttribute("aria-expanded")
                } else {
                    retryClick = true
                }
            }

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

            if (elementsCount > 0) {
                for (let i = 0; i < elementsCount; i++) {
                    const historyElement = historyElements.nth(i)

                    await historyElement.scrollIntoViewIfNeeded()
                    await historyElement.waitFor({ state: "visible", timeout: 10000 })
                    await historyElement.click()

                    const servicoCell = await historyElement.locator(':has-text("Serviço")')
                    const servicoExists = await servicoCell.count()

                    if (servicoExists > 0) {
                        const servicoText = await servicoCell.first().textContent()

                        const regexCentral = /Central de Atendimento/i
                        const regexFornecedor = /Fornecedor SZ/i
                        const regexSistemas = /Sistemas/i
                        const regexEXT = /_EXT$/

                        const hasSistemas = regexSistemas.test(servicoText)
                        const hasCentralAtendimento = regexCentral.test(servicoText)
                        const hasFornecedorSZ = regexFornecedor.test(servicoText)
                        const hasEXT = regexEXT.test(servicoText)

                        if ((hasCentralAtendimento || hasSistemas) && (hasFornecedorSZ || hasEXT)) {
                            console.log(`Found: Central de Atendimento/Sistemas and Fornecedor SZ`)
                            const fullText = await historyElement.locator('xpath=preceding-sibling::div[@id="historyEventItem"][1]').textContent()

                            const dateTimeRegex = /^\d{4}-\d{2}-\d{2}~ \d{2}:\d{2}:\d{2}/
                            const dateTime = fullText.match(dateTimeRegex)[0]

                            const excelDateTime = dateTime.replace("~", " ")

                            console.log("Excel-friendly date and time:", excelDateTime, "\n")

                            if (dateTime) {
                                data[index][newColumnName] = excelDateTime

                                const abertoCell = `A${index + 2}`
                                const caixaSZCell = `L${index + 2}`

                                data[index][
                                    diffColumnName
                                ] = `=INT(${caixaSZCell} - ${abertoCell}) & " dias, " & HORA(${caixaSZCell} - ${abertoCell}) & " horas, " & MINUTO(${caixaSZCell} - ${abertoCell}) & " minutos"`
                            }
                        }
                    }
                }
            }
        }

        const columnOrder = [
            "Aberto",
            "Número",
            "IC Afetado",
            "Empresa Afetada",
            "Aberto por",
            "Descrição resumida",
            "Atribuído a",
            "Encerrado",
            "Estado",
            "SLA de Resolução %",
            "SLA de Resolução Tempo",
            newColumnName,
            diffColumnName,
        ]

        const reorderedData = reorderColumns(data, columnOrder)

        saveToExcelFile(workbook, reorderedData, [newColumnName, diffColumnName], newFilePath)

        console.log("Data saved to file: ", newFilePath)

        index++
    }
    await page.pause()
})()
