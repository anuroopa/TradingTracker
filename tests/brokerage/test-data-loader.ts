import { IDataLoader } from '../../src/brokerage/data-loader-interface';
import { TastyworksHub } from '../../src/brokerage/tastyworks/tastyworks-hub';
import { Portfolio } from '../../src/portfolio/portfolio';
import { OptionsPosition } from '../../src/portfolio/optionsposition';
import { TestData } from './testing/test-data';

export class TestDataLoader implements IDataLoader {
    loadCSVData(portfolio: Portfolio, initialload: boolean = true) {
        let tastyworksHub: TastyworksHub = new TastyworksHub();
        if (initialload) {
            tastyworksHub.loadPortfolioFromCSV(TestData.TASTYWORKS_TEST_PORTFOLIO_CSV, portfolio);
            portfolio.setPositionsNew(false);
        }
        else {
            tastyworksHub.loadPortfolioFromCSV(TestData.TASTYWORKS_CSV_WITH_ROLLS_AND_CLOSINGS, portfolio);
        }
    }

    loadAPIQuoteData(portfolio: Portfolio) {
         //LOAD API DATA
    }
}
