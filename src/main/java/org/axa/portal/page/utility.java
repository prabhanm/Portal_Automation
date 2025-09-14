package org.axa.portal.page;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.Set;
import java.util.TreeMap;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.axa.BC.BC_utility;
import org.axa.framework.CommonFunctions;
import org.axa.framework.Report;
import org.axa.portal.claim.ParameterOfClaimAccidentInformation;
import org.axa.portal.claim.ParameterOfClaimAccidentVehicleInfo;
import org.axa.portal.claim.ParameterOfClaimFirstContact;
import org.axa.portal.validation.ADJ_portal_CommonValidation;

public class utility {

	public static Map<String,ParameterOfHomeAndQuotationPage> quotationPageMap;
	public static Map<String,ParameterOfCurrentInsurancePage> currentInsuranceMap;
	public static Map<String,ParameterOfAboutMainDriverPage> aboutdriverMap;
	public static Map<String,ParameterOfVehicleInformationPage> vehicleInfoMap;
	public static Map<String,ParameterOfPolicyHolderInformationPage> contractorInfoMap;
	public static Map<String,ParameterOfPaymentInformationPage> paymentInfoMap;
	public static Map<String,ParameterOfSuspensionCertificatePage> suspensionCertificateMap;
	public static Map<String,ParameterOfHomeAndQuotationPage> homePageMap;
	public static Map<String,ParameterOfCorporate_PolicyHolderInformationPage> SME_ContractorInfoMap;
	public static Map<String,ParameterOfCreditCardDetails> cardDetailsMap;
	public static Map<String,ParameterOfEmmaLoginScreen> emmaLoginMap;
	public static Map<String,ParameterOfClaimAccidentInformation> claim_AccidentInfoMap;
	public static Map<String,ParameterOfClaimAccidentVehicleInfo> claim_VehicleInfoMap;
	public static Map<String,ParameterOfClaimFirstContact> claim_FirstContactMap;
	public static Properties property;
	public static Map<String,Object[]> resultMap = new TreeMap<>();

	/*
	 * public utility() throws IOException { this.mapItems=
	 * this.getCurrentInsurnacesheetData();
	 * this.homMap=this.getHomeAndQuotationSheetData(); }
	 * 
	 */

	public Properties loadConfigFile(String configFilePath) throws IOException {
		try (FileReader reader = new FileReader(configFilePath)) {
			property = new Properties();
			property.load(reader);
			return property;
		}
	}

	public Map<String,ParameterOfCurrentInsurancePage> getCurrentInsurnaceSheetData(XSSFSheet sheet) throws IOException {

		currentInsuranceMap = new HashMap<>();
		int rowNumber = sheet.getPhysicalNumberOfRows();
		java.util.stream.IntStream.range(1, rowNumber)
			.mapToObj(sheet::getRow)
			.forEach(row -> {
				ParameterOfCurrentInsurancePage parameter = new ParameterOfCurrentInsurancePage();
				parameter.setCurrentInsurnaceNameOfCI(row.getCell(1).getStringCellValue().trim());
				parameter.setUnkownCompanyCheckBoxOfCI(row.getCell(2).getStringCellValue().trim());
				parameter.setContractualConditionOfCI(row.getCell(3).getStringCellValue().trim());
				parameter.setContractPeriodOfCI(row.getCell(4).getStringCellValue().trim());
				parameter.setExpiaryDateOfCI(row.getCell(5).getStringCellValue().trim());
				parameter.setExpairyDateChekboxOfCI(row.getCell(6).getStringCellValue().trim());
				parameter.setGradeTypeOfCI(row.getCell(7).getStringCellValue().trim());
				parameter.setGradeTypeCheckboxOfCI(row.getCell(8).getStringCellValue().trim());
				parameter.setGradeYearOfCI(row.getCell(9).getStringCellValue().trim());
				parameter.setAccidentCofficientOfCI(row.getCell(10).getStringCellValue().trim());
				parameter.setAccidentCofficientCheckboxOfCI(row.getCell(11).getStringCellValue().trim());
				parameter.setAccidentCaseNumberOfCI(row.getCell(12).getStringCellValue().trim());
				parameter.setAccidentTypeOfCI(row.getCell(13).getStringCellValue().trim());
				parameter.setCarInsurnaceQuestionaryOfCI(row.getCell(14).getStringCellValue().trim());
				currentInsuranceMap.put(row.getCell(0).getStringCellValue().trim(), parameter);
			});
		return currentInsuranceMap;

	}

	public Map<String,ParameterOfSuspensionCertificatePage> getSuspensionCertificateSheetData(XSSFSheet sheet) throws IOException {

		suspensionCertificateMap = new HashMap<>();
		int rowNumber = sheet.getPhysicalNumberOfRows();
		java.util.stream.IntStream.range(1, rowNumber)
			.mapToObj(sheet::getRow)
			.forEach(row -> {
				ParameterOfSuspensionCertificatePage parameter = new ParameterOfSuspensionCertificatePage();
				parameter.setSuspensionReason(row.getCell(1).getStringCellValue().trim());
				parameter.setCompanyNameOfSIC(row.getCell(2).getStringCellValue().trim());
				parameter.setUnkownCompanyCheckBoxOfSIC(row.getCell(3).getStringCellValue().trim());
				parameter.setStartdateOfSIC(row.getCell(4).getStringCellValue().trim());
				parameter.setEndDateOfSIC(row.getCell(5).getStringCellValue().trim());
				parameter.setGradeTypeOfSIC(row.getCell(6).getStringCellValue().trim());
				parameter.setGradeTypeCheckboxOfSIC(row.getCell(7).getStringCellValue().trim());
				parameter.setGradeYearOfSIC(row.getCell(8).getStringCellValue().trim());
				parameter.setAccidentCofficientOfSIC(row.getCell(9).getStringCellValue().trim());
				parameter.setAccidentCofficientCheckboxOfSIC(row.getCell(10).getStringCellValue().trim());
				parameter.setAccidentCaseNumberOfSIC(row.getCell(11).getStringCellValue().trim());
				parameter.setAccidentTypeOfSIC(row.getCell(12).getStringCellValue().trim());
				parameter.setCarInsurnaceQuestionaryOfSIC(row.getCell(13).getStringCellValue().trim());
				parameter.setRegistrationDateOfSIC(row.getCell(14).getStringCellValue().trim());
				suspensionCertificateMap.put(row.getCell(0).getStringCellValue().trim(), parameter);
			});
		return suspensionCertificateMap;

	}

	public Map<String,ParameterOfAboutMainDriverPage> getAboutMainDriverSheetData(XSSFSheet sheet) throws IOException {

		aboutdriverMap = new HashMap<>();
		int rowNumber = sheet.getPhysicalNumberOfRows();
		java.util.stream.IntStream.range(1, rowNumber)
			.mapToObj(sheet::getRow)
			.forEach(row -> {
				ParameterOfAboutMainDriverPage parameter = new ParameterOfAboutMainDriverPage();
				parameter.setPolicyHolderIsMainDriver(row.getCell(1).getStringCellValue().trim());
				parameter.setPolicyHolderPrefecture(row.getCell(2).getStringCellValue().trim());
				parameter.setPolicyHolderDOB(row.getCell(3).getStringCellValue().trim());
				parameter.setPolicyHolderLicenceType(row.getCell(4).getStringCellValue().trim());
				parameter.setCarPaxLimitation(row.getCell(5).getStringCellValue().trim());
				parameter.setPolicyHolderAgeRange(row.getCell(6).getStringCellValue().trim());
				parameter.setPolicyPlan(row.getCell(7).getStringCellValue().trim());
				aboutdriverMap.put(row.getCell(0).getStringCellValue().trim(), parameter);
			});
		return aboutdriverMap;

	}

	public Map<String,ParameterOfHomeAndQuotationPage> getQuotationSheetData(XSSFSheet sheet) throws IOException {

		quotationPageMap = new HashMap<>();
		int rowNumber = sheet.getPhysicalNumberOfRows();
		java.util.stream.IntStream.range(1, rowNumber)
			.mapToObj(sheet::getRow)
			.forEach(row -> {
				ParameterOfHomeAndQuotationPage quotationParameter = new ParameterOfHomeAndQuotationPage();
				quotationParameter.setCarRegistrationPeriod(row.getCell(1).getStringCellValue().trim());
				quotationParameter.setUnkownCarRegistrationCheckbox(row.getCell(2).getStringCellValue().trim());
				quotationParameter.setTempCarRegistrationYear(row.getCell(3).getStringCellValue().trim());
				quotationParameter.setBikeDisplacement(row.getCell(4).getStringCellValue().trim());
				quotationParameter.setThreeWheeVechileConfirmation(row.getCell(5).getStringCellValue().trim());
				quotationParameter.setSearchCarByManufacturerCheckbox(row.getCell(6).getStringCellValue().trim());
				quotationParameter.setManufacturerCompanay(row.getCell(7).getStringCellValue().trim());
				quotationParameter.setCarName(row.getCell(8).getStringCellValue().trim());
				quotationParameter.setCar_BikeModel(row.getCell(9).getStringCellValue().trim());
				quotationParameter.setCarBikeMileage(row.getCell(10).getStringCellValue().trim());
				quotationParameter.setVechilePurpose(row.getCell(11).getStringCellValue().trim());
				quotationParameter.setCarUsingwithChildren(row.getCell(12).getStringCellValue().trim());
				quotationParameter.setUnkownChildrenAgeCheckbox(row.getCell(13).getStringCellValue().trim());
				quotationParameter.setChildrenAge(row.getCell(14).getStringCellValue().trim());
				quotationPageMap.put(row.getCell(0).getStringCellValue().trim(), quotationParameter);
			});
		return quotationPageMap;

	}

	public Map<String,ParameterOfVehicleInformationPage> getVehicleInformationSheetData(XSSFSheet sheet) throws IOException {

		vehicleInfoMap = new HashMap<>();
		int rowNumber = sheet.getPhysicalNumberOfRows();
		java.util.stream.IntStream.range(1, rowNumber)
			.mapToObj(sheet::getRow)
			.forEach(row -> {
				ParameterOfVehicleInformationPage vehicleParameter = new ParameterOfVehicleInformationPage();
				vehicleParameter.setGender(row.getCell(1).getStringCellValue().trim());
				vehicleParameter.setCurrentOrSuspendedInsurancePolicyNumber(row.getCell(2).getStringCellValue().trim());
				vehicleParameter.setBranchNumber(row.getCell(3).getStringCellValue().trim());
				vehicleParameter.setCurrentOrsuspendedInsuranceQuestionaries(row.getCell(4).getStringCellValue().trim());
				vehicleParameter.setCarModificationQuestionaries(row.getCell(5).getStringCellValue().trim());
				vehicleParameter.setVehicleChassisNumber(row.getCell(6).getStringCellValue().trim());
				vehicleParameter.setPolicyHolderIsOwner(row.getCell(7).getStringCellValue().trim());
				vehicleParameter.setSecondaryOwnerRelation(row.getCell(8).getStringCellValue().trim());
				vehicleParameter.setSecondaryOwnerLastNameKanji(row.getCell(9).getStringCellValue().trim());
				vehicleParameter.setSecondaryOwnerFirstNameKanji(row.getCell(10).getStringCellValue().trim());
				vehicleParameter.setSecondaryOwnerLastNameKatakana(row.getCell(11).getStringCellValue().trim());
				vehicleParameter.setSecondaryOwnerFirstNameKatakana(row.getCell(12).getStringCellValue().trim());
				vehicleParameter.setLicencePlateKanjiOrPrefecture(row.getCell(13).getStringCellValue().trim());
				vehicleParameter.setLicencePlateNumberOrKatakana(row.getCell(14).getStringCellValue().trim());
				vehicleParameter.setLicencePlateHiragana(row.getCell(15).getStringCellValue().trim());
				vehicleParameter.setLicencePlateSerialNumber(row.getCell(16).getStringCellValue().trim());
				vehicleParameter.setVehicleMileage(row.getCell(17).getStringCellValue().trim());
				vehicleParameter.setVehicleMileageCheckDate(row.getCell(18).getStringCellValue().trim());
				vehicleInfoMap.put(row.getCell(0).getStringCellValue().trim(), vehicleParameter);
			});
		return vehicleInfoMap;

	}

	public Map<String,ParameterOfPolicyHolderInformationPage> getPolicyHolderInformationSheetData(XSSFSheet sheet) throws IOException {

		contractorInfoMap = new HashMap<>();
		int rowNumber = sheet.getPhysicalNumberOfRows();
		java.util.stream.IntStream.range(1, rowNumber)
			.mapToObj(sheet::getRow)
			.forEach(row -> {
				ParameterOfPolicyHolderInformationPage contractorParameter = new ParameterOfPolicyHolderInformationPage();
				contractorParameter.setLastNameKanji(row.getCell(1).getStringCellValue().trim());
				contractorParameter.setFirstNameKanji(row.getCell(2).getStringCellValue().trim());
				contractorParameter.setLastNameFurigana(row.getCell(3).getStringCellValue().trim());
				contractorParameter.setFirstNameFurigana(row.getCell(4).getStringCellValue().trim());
				contractorParameter.setPinCode(row.getCell(5).getStringCellValue().trim());
				contractorParameter.setAddressName(row.getCell(6).getStringCellValue().trim());
				contractorParameter.setDoorNumber(row.getCell(7).getStringCellValue().trim());
				contractorParameter.setBuildingName(row.getCell(8).getStringCellValue().trim());
				contractorParameter.setMobileNumber(row.getCell(9).getStringCellValue().trim());
				contractorParameter.setEmailAddress(row.getCell(10).getStringCellValue().trim());
				contractorParameter.setContractCertificateType(row.getCell(11).getStringCellValue().trim());
				contractorInfoMap.put(row.getCell(0).getStringCellValue().trim(), contractorParameter);
			});
		return contractorInfoMap;

	}

	public Map<String,ParameterOfCorporate_PolicyHolderInformationPage> getSMEPolicyHolderInformationSheetData(XSSFSheet sheet) throws IOException {

		SME_ContractorInfoMap=new HashMap<String,ParameterOfCorporate_PolicyHolderInformationPage>();

		int rowNumber=sheet.getPhysicalNumberOfRows();
		Row row;
		for(int i=1;i<rowNumber;i++) {
			ParameterOfCorporate_PolicyHolderInformationPage SME_contractor=new ParameterOfCorporate_PolicyHolderInformationPage();

			row=sheet.getRow(i);

			SME_contractor.setLegalEntityType(row.getCell(1).getStringCellValue().trim());
			SME_contractor.setCorporateStatusPosition(row.getCell(2).getStringCellValue().trim());
			SME_contractor.setCorporateName(row.getCell(3).getStringCellValue().trim());
			SME_contractor.setCorporateNameKana(row.getCell(4).getStringCellValue().trim());
			SME_contractor.setRepresentativeTitle(row.getCell(5).getStringCellValue().trim());
			SME_contractor.setRepresentativeTitleKana(row.getCell(6).getStringCellValue().trim());
			SME_contractor.setCorporateLastNameKanji(row.getCell(7).getStringCellValue().trim());
			SME_contractor.setCorporateFirstNameKanji(row.getCell(8).getStringCellValue().trim());
			SME_contractor.setCorporateLastNameFurigana(row.getCell(9).getStringCellValue().trim());
			SME_contractor.setCorporateFirstNameFurigana(row.getCell(10).getStringCellValue().trim());
			SME_contractor.setRepresentativePersona(row.getCell(11).getStringCellValue().trim());
			SME_contractor.setDelegateLastNameKanji(row.getCell(12).getStringCellValue().trim());
			SME_contractor.setDelegateFirstNameKanji(row.getCell(13).getStringCellValue().trim());
			SME_contractor.setDelegateLastNameKana(row.getCell(14).getStringCellValue().trim());
			SME_contractor.setDelegateFirstNameKana(row.getCell(15).getStringCellValue().trim());
			SME_contractor.setCorporatePinCode(row.getCell(16).getStringCellValue().trim());
			SME_contractor.setCorporateAddress(row.getCell(17).getStringCellValue().trim());
			SME_contractor.setDoorNumber(row.getCell(18).getStringCellValue().trim());
			SME_contractor.setBuildingName(row.getCell(19).getStringCellValue().trim());
			SME_contractor.setRepresentativePhoneNumber(row.getCell(20).getStringCellValue().trim());
			SME_contractor.setContactNumber(row.getCell(21).getStringCellValue().trim());
			SME_contractor.setEmailAddress(row.getCell(22).getStringCellValue().trim());

			SME_ContractorInfoMap.put(row.getCell(0).getStringCellValue().trim(), SME_contractor);
		}
		return SME_ContractorInfoMap;

	}

	public Map<String,ParameterOfPaymentInformationPage> getPaymentInformationSheetData(XSSFSheet sheet) throws IOException {

		paymentInfoMap=new HashMap<String,ParameterOfPaymentInformationPage>();

		int rowNumber=sheet.getPhysicalNumberOfRows();
		Row row;
		for(int i=1;i<rowNumber;i++) {
			ParameterOfPaymentInformationPage payment=new ParameterOfPaymentInformationPage();

			row=sheet.getRow(i);

			payment.setPaymentMode(row.getCell(1).getStringCellValue().trim());
			payment.setCardOrStoreType(row.getCell(2).getStringCellValue().trim());

			paymentInfoMap.put(row.getCell(0).getStringCellValue().trim(), payment);
		}
		return paymentInfoMap;

	}

	public Map<String,ParameterOfCreditCardDetails> getCreditCardDetailsSheetData(XSSFSheet sheet) throws IOException {

		cardDetailsMap=new HashMap<String,ParameterOfCreditCardDetails>();

		int rowNumber=sheet.getPhysicalNumberOfRows();
		Row row;
		for(int i=1;i<rowNumber;i++) {
			ParameterOfCreditCardDetails cardDetails=new ParameterOfCreditCardDetails();

			row=sheet.getRow(i);

			cardDetails.setCardNumber(row.getCell(1).getStringCellValue().trim());
			cardDetails.setCardExpiaryDate(row.getCell(2).getStringCellValue().trim());
			cardDetails.setCvvNumber(row.getCell(3).getStringCellValue().trim());

			cardDetailsMap.put(row.getCell(0).getStringCellValue().trim(), cardDetails);
		}
		return cardDetailsMap;

	}

	public Map<String,ParameterOfEmmaLoginScreen> getEmmaLogin(XSSFSheet sheet) throws IOException {

		emmaLoginMap=new HashMap<String,ParameterOfEmmaLoginScreen>();

		int rowNumber=sheet.getPhysicalNumberOfRows();
		Row row;
		for(int i=1;i<rowNumber;i++) {
			ParameterOfEmmaLoginScreen loginDetails=new ParameterOfEmmaLoginScreen();

			row=sheet.getRow(i);

			loginDetails.setLoginID(row.getCell(1).getStringCellValue().trim());
			loginDetails.setPassword(row.getCell(2).getStringCellValue().trim());
			loginDetails.setRecoveryType(row.getCell(3).getStringCellValue().trim());
			loginDetails.setPolicyNumber(row.getCell(4).getStringCellValue().trim());
			loginDetails.setLastName(row.getCell(5).getStringCellValue().trim());
			loginDetails.setFirstName(row.getCell(6).getStringCellValue().trim());
			loginDetails.setMobileNumber(row.getCell(7).getStringCellValue().trim());

			emmaLoginMap.put(row.getCell(0).getStringCellValue().trim(), loginDetails);
		}
		return emmaLoginMap;

	}
	
	public Map<String,ParameterOfClaimAccidentInformation> getClaimAccidentInformation(XSSFSheet sheet) throws IOException {

		claim_AccidentInfoMap=new HashMap<String,ParameterOfClaimAccidentInformation>();

		int rowNumber=sheet.getPhysicalNumberOfRows();
		Row row;
		for(int i=1;i<rowNumber;i++) {
			ParameterOfClaimAccidentInformation accidentDetails=new ParameterOfClaimAccidentInformation();

			row=sheet.getRow(i);

			accidentDetails.setDateOfAccident(row.getCell(1).getStringCellValue().trim());
			accidentDetails.setTimeOfAccident(row.getCell(2).getStringCellValue().trim());
			accidentDetails.setLocationOfAccident(row.getCell(3).getStringCellValue().trim());
			accidentDetails.setAddressOfAccident(row.getCell(4).getStringCellValue().trim());
			accidentDetails.setTypeOfAccident(row.getCell(5).getStringCellValue().trim());
			accidentDetails.setInjuryStatus(row.getCell(6).getStringCellValue().trim());
			accidentDetails.setAccidentExplanation1(row.getCell(7).getStringCellValue().trim());
			accidentDetails.setAccidentExplanation2(row.getCell(8).getStringCellValue().trim());
			accidentDetails.setTrafficLight(row.getCell(9).getStringCellValue().trim());
			accidentDetails.setTrafficLightMovement(row.getCell(10).getStringCellValue().trim());
			accidentDetails.setAccidentCircumstance(row.getCell(11).getStringCellValue().trim());
			accidentDetails.setRemarks(row.getCell(12).getStringCellValue().trim());

			claim_AccidentInfoMap.put(row.getCell(0).getStringCellValue().trim(), accidentDetails);
		}
		return claim_AccidentInfoMap;

	}
	
	public Map<String,ParameterOfClaimAccidentVehicleInfo> getClaimAccidentVehicleInformation(XSSFSheet sheet) throws IOException {

		claim_VehicleInfoMap=new HashMap<String,ParameterOfClaimAccidentVehicleInfo>();

		int rowNumber=sheet.getPhysicalNumberOfRows();
		Row row;
		for(int i=1;i<rowNumber;i++) {
			ParameterOfClaimAccidentVehicleInfo vehicleDetails=new ParameterOfClaimAccidentVehicleInfo();

			row=sheet.getRow(i);

			
			vehicleDetails.setVehicleType(row.getCell(1).getStringCellValue().trim());
			vehicleDetails.setVehicleRepairStatus(row.getCell(2).getStringCellValue().trim());
			vehicleDetails.setRepairShopName(row.getCell(3).getStringCellValue().trim());
			vehicleDetails.setRepairShopPhoneNumber(row.getCell(4).getStringCellValue().trim());
			vehicleDetails.setVehicleDriver(row.getCell(5).getStringCellValue().trim());
			vehicleDetails.setDriverLastNameKanji(row.getCell(6).getStringCellValue().trim());
			vehicleDetails.setDriverFirstNameKanji(row.getCell(7).getStringCellValue().trim());
			vehicleDetails.setDriverLastNameKana(row.getCell(8).getStringCellValue().trim());
			vehicleDetails.setDriverFirstNameKana(row.getCell(9).getStringCellValue().trim());
			vehicleDetails.setLicencePhotoPath(row.getCell(10).getStringCellValue().trim());
			vehicleDetails.setLicenceColour(row.getCell(11).getStringCellValue().trim());
			vehicleDetails.setLicenceExpiaryDate(row.getCell(12).getStringCellValue().trim());
			vehicleDetails.setLicenceNumber(row.getCell(13).getStringCellValue().trim());
			vehicleDetails.setDamagedCarPhotoPath(row.getCell(13).getStringCellValue().trim());

			claim_VehicleInfoMap.put(row.getCell(0).getStringCellValue().trim(), vehicleDetails);
		}
		return claim_VehicleInfoMap;

	}
	
	public Map<String,ParameterOfClaimFirstContact> getClaimFirstContactInformation(XSSFSheet sheet) throws IOException {

		claim_FirstContactMap=new HashMap<String,ParameterOfClaimFirstContact>();

		int rowNumber=sheet.getPhysicalNumberOfRows();
		Row row;
		for(int i=1;i<rowNumber;i++) {
			ParameterOfClaimFirstContact firstContact=new ParameterOfClaimFirstContact();

			row=sheet.getRow(i);

			
			firstContact.setContactAddress(row.getCell(1).getStringCellValue().trim());
			firstContact.setPhoneNumberType(row.getCell(2).getStringCellValue().trim());
			firstContact.setAddressiLastNameKanji(row.getCell(3).getStringCellValue().trim());
			firstContact.setAddressiFirstNameKanji(row.getCell(4).getStringCellValue().trim());
			firstContact.setAddressiLastNameKana(row.getCell(5).getStringCellValue().trim());
			firstContact.setAddressiFirstNameKana(row.getCell(6).getStringCellValue().trim());
			firstContact.setAddressiMobileNumber(row.getCell(7).getStringCellValue().trim());
			firstContact.setContactTime(row.getCell(8).getStringCellValue().trim());
			firstContact.setNotificationMethod(row.getCell(9).getStringCellValue().trim());
			firstContact.setNotificationPhoneNumber(row.getCell(10).getStringCellValue().trim());
			firstContact.setNotificationRemarks(row.getCell(11).getStringCellValue().trim());

			claim_FirstContactMap.put(row.getCell(0).getStringCellValue().trim(), firstContact);
		}
		return claim_FirstContactMap;

	}

	public List<String> getExecutionDataSheet(String summryDataSheet) throws IOException {

	File file = new File(System.getProperty("user.dir") + "\\Data_Input\\Portal\\" + summryDataSheet);
	try (FileInputStream inputStream = new FileInputStream(file);
	     XSSFWorkbook workbook = new XSSFWorkbook(inputStream)) {
	    XSSFSheet sheet = workbook.getSheet("Execution_Summary");
	    int rowNumber = sheet.getPhysicalNumberOfRows();
	    return java.util.stream.IntStream.range(1, rowNumber)
		.mapToObj(sheet::getRow)
		.filter(row -> row.getCell(1).getNumericCellValue() == 1)
		.map(row -> row.getCell(2).getStringCellValue())
		.collect(java.util.stream.Collectors.toList());
	}
	}

	public void methodToInokeFunction() throws IOException, ClassNotFoundException, NoSuchMethodException, SecurityException, IllegalAccessException, IllegalArgumentException, InvocationTargetException, InstantiationException {
		loadConfigFile(".//portalConfiguration.properties");
		System.out.println("I am here");

		List<ParameterOfHomeAndQuotationPage> list = new ArrayList<>();
		homePageMap = new HashMap<>();

		getExecutionDataSheet(property.getProperty("InputDatasheet")).forEach(fileName -> {
			File file = new File(System.getProperty("user.dir") + "\\Data_Input\\Portal\\" + fileName);
			try (FileInputStream inputStream = new FileInputStream(file);
				 XSSFWorkbook workbook = new XSSFWorkbook(inputStream)) {
				int sheetCount = workbook.getNumberOfSheets();
				System.out.println("I am here2");
				java.util.stream.IntStream.range(0, sheetCount)
					.mapToObj(workbook::getSheetName)
					.forEach(sheetName -> {
						System.out.println(sheetName);
						XSSFSheet sheet = workbook.getSheet(sheetName);
						switch (sheetName) {
							case "Final summary":
								int rowNumber = sheet.getPhysicalNumberOfRows();
								java.util.stream.IntStream.range(1, rowNumber)
									.mapToObj(sheet::getRow)
									.forEach(row -> {
										ParameterOfHomeAndQuotationPage parameter = new ParameterOfHomeAndQuotationPage();
										System.out.println("I am here3");
										parameter.setTestCaseID(row.getCell(0).getStringCellValue().trim());
										parameter.setBrowserView(row.getCell(1).getStringCellValue().trim());
										parameter.setExecutionFlowType(row.getCell(2).getStringCellValue().trim());
										parameter.setTestRunstatus(row.getCell(3).getNumericCellValue());
										parameter.setInsurnaceFlowType(row.getCell(4).getStringCellValue().trim());
										parameter.setMemberType(row.getCell(5).getStringCellValue().trim());
										parameter.setInsurnaceProductType(row.getCell(6).getStringCellValue().trim());
										parameter.setInsurancePurchaseType(row.getCell(7).getStringCellValue().trim());
										parameter.setPurchaseInsurnaceInceptionDate(row.getCell(8).getStringCellValue().trim());
										parameter.setUnknownInceptionDateCheckbox(row.getCell(9).getStringCellValue().trim());
										parameter.setTestCaseOverview(row.getCell(10).getStringCellValue().trim());
										list.add(parameter);
										homePageMap.put(row.getCell(0).getStringCellValue().trim(), parameter);
									});
								break;
							case "Emma Login":
							try {
								getEmmaLogin(sheet);
							} catch (IOException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
								break;
							case "Current Insurance screen":
							try {
								getCurrentInsurnaceSheetData(sheet);
							} catch (IOException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
								break;
							case "Suspension Certificate Screen":
							try {
								getSuspensionCertificateSheetData(sheet);
							} catch (IOException e1) {
								// TODO Auto-generated catch block
								e1.printStackTrace();
							}
								break;
							case "Quotation Page":
							try {
								getQuotationSheetData(sheet);
							} catch (IOException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
								break;
							case "About Main Driver":
							try {
								getAboutMainDriverSheetData(sheet);
							} catch (IOException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
								break;
							case "Vehicle Information":
							try {
								getVehicleInformationSheetData(sheet);
							} catch (IOException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
								break;
							case "Corporate Contractor Info":
							try {
								getSMEPolicyHolderInformationSheetData(sheet);
							} catch (IOException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
								break;
							case "Contractor Information":
							try {
								getPolicyHolderInformationSheetData(sheet);
							} catch (IOException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
								break;
							case "Payment Information":
							try {
								getPaymentInformationSheetData(sheet);
							} catch (IOException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
								break;
							case "Credit Card Details":
							try {
								getCreditCardDetailsSheetData(sheet);
							} catch (IOException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
								break;
							case "Accident Information":
							try {
								getClaimAccidentInformation(sheet);
							} catch (IOException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
								break;
							case "Accident_Vehicle Info":
							try {
								getClaimAccidentVehicleInformation(sheet);
							} catch (IOException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
								break;
							case "First Contract":
							try {
								getClaimFirstContactInformation(sheet);
							} catch (IOException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
								break;
							default:
								break;
						}
					});
			} catch (Exception e) {
				// Consider logging exception
			}
		});
		Report.reportInitilization("Portal", utility.property.getProperty("Environment"));
		list.stream().filter(data -> (int) data.getTestRunstatus() == 1).forEach(data -> {
			ADJ_portal_currentInsuranceScreen.INSURANCE_START_DATE = null;
			ADJ_portal_homePage.PURCHASE_FLOW_TYPE = null;
			ADJ_portal_homePage.quotationNumber = null;
			ADJ_portal_CommonValidation.quotationNumberOfVehiclePage = null;
			ADJ_portal_aboutMainDriverAndPolicyPlanScreen.POLICY_PLAN_AMOUNT = null;
			ADJ_portal_quotationScreen.CAR_MANUFACTURERE_AFTER_SELECTION = null;
			ADJ_portal_quotationScreen.CAR_MODEL_NUMBER_AFTER_SELECTION = null;
			ADJ_portal_quotationScreen.CAR_NAME_AFTER_SELECTION = null;
			ADJ_portal_ContractConfirmationScreen.dataList = new ArrayList<>();
			ADJ_portal_paymentInformationScreen.FINAL_PREMIUM_AMOUNT = null;
			System.out.println("---------------Test Case execution is started for case: " + data.getTestCaseID() + "---------------");
			Report.startTest(data.getTestCaseID() + ":" + data.getTestCaseOverview());
			String testCaseName = data.getExecutionFlowType().equals("End2End") ? data.getExecutionFlowType() : data.getTestCaseID();
			try {
				Class<?> c = Class.forName("com.axa.test.ADJ_portal_testCaseScript");
				Object obj = c.getDeclaredConstructor().newInstance();
				Method method = obj.getClass().getMethod(testCaseName, ParameterOfHomeAndQuotationPage.class);
				method.invoke(obj, data);
			} catch (Exception e) {
				// Consider logging exception
			}
		});
		Report.endReport();
		createResultDataSheet();
		System.out.println("-----------------Automation Execution is completed------------------");
	}

	public static void storeResultData(String testCaseID,String status,String policyNumber) {
		try {
			Object[] arr=new Object[ADJ_portal_ContractConfirmationScreen.dataList.size()+2];
			arr[0]=status;
			arr[1]=policyNumber;

			for(int i=0;i<ADJ_portal_ContractConfirmationScreen.dataList.size();i++) {
				arr[2+i]=ADJ_portal_ContractConfirmationScreen.dataList.get(i);
			}
			System.out.println(arr[13]);
			System.out.println(arr[15]);
			arr[15]=ADJ_portal_paymentInformationScreen.FINAL_PREMIUM_AMOUNT;
			System.out.println(arr[15]);
			resultMap.put(testCaseID, arr);
			//resultMap.put(testCaseID, new Object[] {status,policyNumber,ADJ_portal_ContractConfirmationScreen.dataList.get(0),ADJ_portal_ContractConfirmationScreen.dataList.get(1),ADJ_portal_ContractConfirmationScreen.dataList.get(2),ADJ_portal_ContractConfirmationScreen.dataList.get(3),ADJ_portal_ContractConfirmationScreen.dataList.get(4),ADJ_portal_ContractConfirmationScreen.dataList.get(5)});
		}catch(Exception e) {

		}
	}

	public void createResultDataSheet() throws IOException {
		try (XSSFWorkbook workbook = new XSSFWorkbook()) {
			XSSFSheet sheet = workbook.createSheet("Result");
			Set<String> keyid = resultMap.keySet();
			XSSFRow row = sheet.createRow(0);
			row.createCell(0).setCellValue("TEST_ID");
			row.createCell(1).setCellValue("STATUS");
			row.createCell(2).setCellValue("TRANSACTION_NUMBER");
			row.createCell(3).setCellValue("EXECUTION_FLOW");
			row.createCell(4).setCellValue("EFFECTIVE_DATE");
			row.createCell(5).setCellValue("INCEPTION_DATE");
			row.createCell(6).setCellValue("MATURIRY_DATE");
			row.createCell(7).setCellValue("VEHICLE_TYPE");
			row.createCell(8).setCellValue("VEHICLE_NAME");
			row.createCell(9).setCellValue("VEHICLE_MODEL");
			row.createCell(10).setCellValue("POLICYHOLDER_NAME");
			row.createCell(11).setCellValue("POLICYHOLDER_NAME_KANA");
			row.createCell(12).setCellValue("CORPOARTE_NAME");
			row.createCell(13).setCellValue("CORPOARTE_NAME_KANA");
			row.createCell(14).setCellValue("POSTLE_CODE");
			row.createCell(15).setCellValue("ADDRESS");
			row.createCell(16).setCellValue("PREMENIUM_AMOUNT");
			row.createCell(17).setCellValue("COVERAGE_TYPE");
			row.createCell(18).setCellValue("COVERAGE_AMOUNT");
			row.createCell(19).setCellValue("PAYMENT_MODE");
			row.createCell(20).setCellValue("PHONE_NUMBER");
			row.createCell(21).setCellValue("EMAIL");
			row.createCell(22).setCellValue("DOB");
			row.createCell(23).setCellValue("CORPORATE_POSITION");
			row.createCell(24).setCellValue("CORPORATE_TITLE");
			row.createCell(25).setCellValue("CORPORATE_TITLE_KANA");
			row.createCell(26).setCellValue("ACCOUNT_NUMBER");
			int rowid = 1;
			for (String key : keyid) {
				row = sheet.createRow(rowid++);
				Object[] objectArr = resultMap.get(key);
				row.createCell(0).setCellValue(key);
				int cellid = 1;
				for (Object obj : objectArr) {
					row.createCell(cellid++).setCellValue(obj != null ? obj.toString() : "");
				}
			}
			try (FileOutputStream out = new FileOutputStream(new File(System.getProperty("user.dir") + Report.CAPTURE_PATH + "\\Result.xlsx"))) {
				workbook.write(out);
			}
		}


	}}
