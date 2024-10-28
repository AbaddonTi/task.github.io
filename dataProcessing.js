// dataProcessing.js

async function fetchUniqueBanks(records) {
    try {
        const bankColumn = 'Банк';

        if (!records || records.length === 0) {
            
            return [];
        }

        const banks = records.map(record => record[bankColumn]);

        const uniqueBanks = [...new Set(banks)].filter(bank => bank);

        return uniqueBanks;
    } catch (error) {
        console.error('Ошибка при получении банков:', error);
        return [];
    }
}

async function calculateMetrics(records, banks) {
    const paidColumn = 'Выплачено';
    try {
        const priceColumn = 'Цена';
        const bankColumn = 'Банк';
        const profitColumn = 'Профит';
        const statusColumn = 'Статус';

        const bankPrices = {};
        const bankCardCounts = {};
        const bankProfits = {};
        const bankProblemCounts = {};
        const bankBlockCounts = {};

        banks.forEach(bank => {
            bankPrices[bank] = [];
            bankCardCounts[bank] = 0;
            bankProfits[bank] = 0;
            bankProblemCounts[bank] = 0;
            bankBlockCounts[bank] = 0;
        });

        let totalBankCount = 0;
        let totalPriceSum = 0;
        let totalProfitSum = 0;
        let totalProblemCount = 0;
        let totalBlockCount = 0;
        let totalPaidForProblem = 0;
        let totalPaidForBlock = 0;

        records.forEach(record => {
            const bank = record[bankColumn];
            const price = record[priceColumn];
            const profit = record[profitColumn];
            const status = record[statusColumn];
            if (bank && bank.trim() !== '') {
                if (bankPrices[bank] !== undefined) {
                    bankPrices[bank].push(price || 0);
                    bankCardCounts[bank] += 1;
                    bankProfits[bank] += profit || 0;
                    totalBankCount += 1;
                    totalPriceSum += price || 0;
                    totalProfitSum += profit || 0;
                    if (status && status.toLowerCase() === 'проблема') {
                        bankProblemCounts[bank] += 1;
                        totalProblemCount += 1;
                        totalPaidForProblem += record[paidColumn] || 0;
                    }
                    if (status && status.toLowerCase() === 'блокировка') {
                        bankBlockCounts[bank] += 1;
                        totalBlockCount += 1;
                        totalPaidForBlock += record[paidColumn] || 0;
                    }
                }
            }
        });

        let averagePrice = 0;
        if (totalBankCount > 0) {
            averagePrice = totalPriceSum / totalBankCount;
        }

        let averageProfit = 0;
        if (totalBankCount > 0) {
            averageProfit = totalProfitSum / totalBankCount;
        }

        let problemPercentage = 0;
        if (totalBankCount > 0) {
            problemPercentage = (totalProblemCount / totalBankCount) * 100;
        }

        const totalProcessingCol = 1 + banks.length;
        setCellText(6, totalProcessingCol, averagePrice.toFixed(2));
        setCellText(7, totalProcessingCol, totalBankCount);
        setCellText(8, totalProcessingCol, averageProfit.toFixed(2));
        setCellText(10, totalProcessingCol, totalProfitSum.toFixed(2));
        setCellText(12, totalProcessingCol, (problemPercentage / 100).toFixed(4));
        let blockPercentage = 0;
        if (totalBankCount > 0) {
            blockPercentage = (totalBlockCount / totalBankCount) * 100;
        }
        setCellText(16, totalProcessingCol, (blockPercentage / 100).toFixed(4));
        setCellText(14, totalProcessingCol, totalProblemCount);
        setCellText(18, totalProcessingCol, totalBlockCount);
        const totalProcessingColLetter = String.fromCharCode(65 + totalProcessingCol);
        setCellText(20, totalProcessingCol, `=${totalProcessingColLetter}8*${totalProcessingColLetter}7`);
        
        let sumFormula = "=";
        for (let col = 1; col <= banks.length; col++) {
            const colLetter = String.fromCharCode(65 + col);
            sumFormula += `${colLetter}23`;
            if (col < banks.length) {
                sumFormula += "+";
            }
        }
        setCellText(22, totalProcessingCol, sumFormula);

        setCellText(13, totalProcessingCol, totalPaidForProblem.toFixed(2));
        setCellText(17, totalProcessingCol, totalPaidForBlock.toFixed(2));

        let totalCostsSumFormula = "=";
        for (let col = 1; col <= banks.length; col++) {
            const colLetter = String.fromCharCode(65 + col);
            totalCostsSumFormula += `${colLetter}25`;
            if (col < banks.length) {
                totalCostsSumFormula += "+";
            }
        }
        setCellText(24, totalProcessingCol, totalCostsSumFormula);

        let db1SumFormula = "=";
        for (let col = 1; col <= banks.length; col++) {
            const colLetter = String.fromCharCode(65 + col);
            db1SumFormula += `${colLetter}26`;
            if (col < banks.length) {
                db1SumFormula += "+";
            }
        }
        setCellText(25, totalProcessingCol, db1SumFormula);

        const averagePrices = banks.map(bank => {
            const prices = bankPrices[bank];
            if (prices.length > 0) {
                const total = prices.reduce((sum, price) => sum + price, 0);
                return total / prices.length;
            }
            return 0;
        });

        const averageProfits = banks.map(bank => {
            if (bankCardCounts[bank] > 0) {
                return bankProfits[bank] / bankCardCounts[bank];
            }
            return 0;
        });

        const problemPercentages = banks.map(bank => {
            if (bankCardCounts[bank] > 0) {
                return (bankProblemCounts[bank] / bankCardCounts[bank]) * 100;
            }
            return 0;
        });

        const blockPercentages = banks.map(bank => {
            if (bankCardCounts[bank] > 0) {
                return (bankBlockCounts[bank] / bankCardCounts[bank]) * 100;
            }
            return 0;
        });

        return { averagePrices, bankCardCounts, averageProfits, problemPercentages, blockPercentages, bankProblemCounts, bankBlockCounts };
    } catch (error) {
        console.error('Ошибка при вычислении метрик:', error);
        return { averagePrices: [], bankCardCounts: {}, averageProfits: [], problemPercentages: [], blockPercentages: [], bankProblemCounts: {}, bankBlockCounts: {} };
    }
}
