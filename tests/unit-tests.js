/*
 * Excel Data Cleaner Add-in
 * Unit Tests for Core Functionality
 */

import { expect } from 'chai';
import sinon from 'sinon';
import { JSDOM } from 'jsdom';

// Import functions to test
import { convertToProperCase, convertToSentenceCase } from '../features/basic/case.js';
import { escapeRegExp } from '../features/basic/replace.js';
import { tryParseDate, formatDate } from '../features/premium/dates.js';
import { detectDateFormat } from '../utils/enhanced-date-detection.js';
import { applyCustomTrim, applyCustomCaseConversion } from '../utils/customizable-settings.js';
import { processRangeInChunks } from '../utils/performance-optimization.js';

// Setup DOM environment for tests
const jsdom = new JSDOM('<!doctype html><html><body></body></html>');
global.window = jsdom.window;
global.document = jsdom.window.document;

describe('Excel Data Cleaner - Unit Tests', () => {
  // Case Conversion Tests
  describe('Case Conversion Functions', () => {
    it('should convert text to proper case', () => {
      expect(convertToProperCase('hello world')).to.equal('Hello World');
      expect(convertToProperCase('SALES REPORT')).to.equal('Sales Report');
      expect(convertToProperCase('mixed CASE text')).to.equal('Mixed Case Text');
    });
    
    it('should handle special characters in proper case conversion', () => {
      expect(convertToProperCase('hello-world')).to.equal('Hello-World');
      expect(convertToProperCase('john smith jr.')).to.equal('John Smith Jr.');
      expect(convertToProperCase('"quoted" text')).to.equal('"Quoted" Text');
    });
    
    it('should convert text to sentence case', () => {
      expect(convertToSentenceCase('hello world.')).to.equal('Hello world.');
      expect(convertToSentenceCase('hello. world. test.')).to.equal('Hello. World. Test.');
      expect(convertToSentenceCase('SENTENCE. another sentence.')).to.equal('Sentence. Another sentence.');
    });
    
    it('should apply custom case conversion with settings', () => {
      // Mock settingsManager
      global.settingsManager = {
        getSetting: sinon.stub()
      };
      
      // Test UPPER case
      global.settingsManager.getSetting.withArgs('caseConversion.preserveCase', []).returns(['USA', 'iPhone']);
      global.settingsManager.getSetting.withArgs('caseConversion.respectAcronyms', true).returns(true);
      expect(applyCustomCaseConversion('hello usa and iphone', 'UPPER')).to.equal('HELLO USA AND IPHONE');
      
      // Test LOWER case
      expect(applyCustomCaseConversion('HELLO USA AND IPHONE', 'LOWER')).to.equal('hello usa and iphone');
      
      // Test PROPER case with preserved words
      const result = applyCustomCaseConversion('hello usa and iphone world', 'PROPER');
      expect(result).to.equal('Hello USA and iPhone World');
    });
  });
  
  // String Manipulation Tests
  describe('String Manipulation Functions', () => {
    it('should escape regular expression special characters', () => {
      expect(escapeRegExp('hello.world')).to.equal('hello\\.world');
      expect(escapeRegExp('(test)')).to.equal('\\(test\\)');
      expect(escapeRegExp('a+b*c?')).to.equal('a\\+b\\*c\\?');
    });
    
    it('should trim spaces according to settings', () => {
      // Mock settingsManager
      global.settingsManager = {
        getSetting: sinon.stub()
      };
      
      // Test with all trim options enabled
      global.settingsManager.getSetting.withArgs('trimSpaces.trimLeft', true).returns(true);
      global.settingsManager.getSetting.withArgs('trimSpaces.trimRight', true).returns(true);
      global.settingsManager.getSetting.withArgs('trimSpaces.reduceDuplicateSpaces', true).returns(true);
      
      expect(applyCustomTrim('  hello  world  ')).to.equal('hello world');
      expect(applyCustomTrim('multiple   spaces')).to.equal('multiple spaces');
      
      // Test with only left trim enabled
      global.settingsManager.getSetting.withArgs('trimSpaces.trimLeft', true).returns(true);
      global.settingsManager.getSetting.withArgs('trimSpaces.trimRight', true).returns(false);
      global.settingsManager.getSetting.withArgs('trimSpaces.reduceDuplicateSpaces', true).returns(false);
      
      expect(applyCustomTrim('  hello  ')).to.equal('hello  ');
      
      // Test with only right trim enabled
      global.settingsManager.getSetting.withArgs('trimSpaces.trimLeft', true).returns(false);
      global.settingsManager.getSetting.withArgs('trimSpaces.trimRight', true).returns(true);
      global.settingsManager.getSetting.withArgs('trimSpaces.reduceDuplicateSpaces', true).returns(false);
      
      expect(applyCustomTrim('  hello  ')).to.equal('  hello');
      
      // Test with only duplicate space reduction enabled
      global.settingsManager.getSetting.withArgs('trimSpaces.trimLeft', true).returns(false);
      global.settingsManager.getSetting.withArgs('trimSpaces.trimRight', true).returns(false);
      global.settingsManager.getSetting.withArgs('trimSpaces.reduceDuplicateSpaces', true).returns(true);
      
      expect(applyCustomTrim('  hello   world  ')).to.equal('  hello world  ');
    });
  });
  
  // Date Handling Tests
  describe('Date Handling Functions', () => {
    it('should try to parse various date formats', () => {
      // Date object
      const dateObj = new Date(2023, 0, 15); // Jan 15, 2023
      expect(tryParseDate(dateObj)).to.be.an.instanceof(Date);
      expect(tryParseDate(dateObj).getFullYear()).to.equal(2023);
      
      // Excel date serial number (Jan 15, 2023)
      const result1 = tryParseDate(44941);
      expect(result1).to.be.an.instanceof(Date);
      
      // String date formats
      expect(tryParseDate('01/15/2023')).to.be.an.instanceof(Date);
      expect(tryParseDate('2023-01-15')).to.be.an.instanceof(Date);
      expect(tryParseDate('15-Jan-2023')).to.be.an.instanceof(Date);
      
      // Invalid dates
      expect(tryParseDate('not a date')).to.be.null;
      expect(tryParseDate('13/40/2023')).to.be.null; // Invalid month/day
    });
    
    it('should format dates according to specified format', () => {
      const date = new Date(2023, 0, 15); // Jan 15, 2023
      
      expect(formatDate(date, 'MM/DD/YYYY')).to.equal('01/15/2023');
      expect(formatDate(date, 'DD/MM/YYYY')).to.equal('15/01/2023');
      expect(formatDate(date, 'YYYY-MM-DD')).to.equal('2023-01-15');
    });
    
    it('should detect date formats correctly', () => {
      // ISO format
      const isoResult = detectDateFormat('2023-01-15');
      expect(isoResult).to.not.be.null;
      expect(isoResult.format).to.equal('yyyy-mm-dd');
      
      // US format
      const usResult = detectDateFormat('01/15/2023');
      expect(usResult).to.not.be.null;
      expect(usResult.format).to.equal('mm/dd/yyyy');
      
      // European format
      const euResult = detectDateFormat('15/01/2023');
      expect(euResult).to.not.be.null;
      expect(euResult.format).to.equal('dd/mm/yyyy');
      
      // Month name format
      const monthNameResult = detectDateFormat('15-Jan-2023');
      expect(monthNameResult).to.not.be.null;
      expect(monthNameResult.format).to.equal('dd-mmm-yyyy');
      
      // Not a date
      expect(detectDateFormat('not a date')).to.be.null;
    });
  });
  
  // Performance Optimization Tests
  describe('Performance Optimization', () => {
    it('should process range in chunks', async () => {
      // Mock Excel.RequestContext and Range
      const mockContext = {
        sync: sinon.stub().resolves()
      };
      
      const mockRange = {
        rowCount: 2500,
        columnCount: 5,
        address: 'A1:E2500',
        worksheet: {
          getRange: sinon.stub().returns({
            getOffsetRange: sinon.stub().returns({
              getResizedRange: sinon.stub().returns({})
            })
          })
        },
        load: sinon.stub()
      };
      
      // Mock process chunk function
      const processChunk = sinon.stub().resolves({ changedItems: 10 });
      const onProgress = sinon.spy();
      
      // Process range in chunks
      const result = await processRangeInChunks(mockContext, mockRange, processChunk, {
        chunkSize: 1000,
        onProgress
      });
      
      // Verify results
      expect(processChunk.callCount).to.equal(3); // Should call 3 times for 2500 rows with chunk size 1000
      expect(onProgress.callCount).to.equal(3);
      expect(result.processedRows).to.equal(2500);
      expect(result.changedItems).to.equal(30); // 10 per chunk * 3 chunks
    });
  });
  
  // Excel API Integration Mocks
  describe('Excel API Integration', () => {
    beforeEach(() => {
      // Mock Excel object
      global.Excel = {
        run: async (callback) => {
          // Mock context for Excel.run
          const context = {
            workbook: {
              worksheets: {
                getActiveWorksheet: () => ({
                  getRange: () => ({})
                })
              },
              getSelectedRange: () => ({
                load: () => {},
                values: [['  Test  '], ['UPPER'], ['lower']],
                getCell: () => ({
                  numberFormat: ''
                })
              })
            },
            sync: async () => {}
          };
          
          return await callback(context);
        }
      };
    });
    
    afterEach(() => {
      delete global.Excel;
    });
    
    it('should test a basic Excel integration', async () => {
      // Create a sample operation function
      const testOperation = async () => {
        let result = false;
        
        await Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.load('values');
          await context.sync();
          
          if (range.values[0][0] === '  Test  ') {
            result = true;
          }
        });
        
        return result;
      };
      
      const result = await testOperation();
      expect(result).to.be.true;
    });
  });
});