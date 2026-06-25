import { Asset } from "expo-asset";
import * as FileSystem from "expo-file-system/legacy";
import * as Sharing from "expo-sharing";
import { useState } from "react";
import {
  Alert,
  Platform,
  ScrollView,
  StyleSheet,
  Text,
  TextInput,
  TouchableOpacity,
  View
} from "react-native";
import * as XLSX from "xlsx";

type SectionKey =
  | "customer"
  | "equipment"
  | "accessories"
  | "materials"
  | "services"
  | "scope"
  | "warranty"
  | "payment";

type GridValue = "" | "Y" | "N" | "CUSTOM";

type MaterialItem = {
  label: string;
  cells: string[];
  values: GridValue[];
  customValues: string[];
};

type FieldProps = {
  label: string;
  value: string;
  onChangeText: (text: string) => void;
  placeholder?: string;
  keyboardType?: "default" | "numeric" | "phone-pad";
  multiline?: boolean;
};

function Field({
  label,
  value,
  onChangeText,
  placeholder,
  keyboardType,
  multiline = false,
}: FieldProps) {
  return (
    <View style={fieldStyles.wrap}>
      <Text selectable={false} style={fieldStyles.label}>
        {label}
      </Text>

      <TextInput
        style={[fieldStyles.input, multiline && fieldStyles.textArea]}
        value={value}
        onChangeText={onChangeText}
        placeholder={placeholder || label}
        placeholderTextColor="#7f8b96"
        keyboardType={keyboardType || "default"}
        multiline={multiline}
        textAlignVertical={multiline ? "top" : "center"}
        autoCorrect={false}
        spellCheck={false}
        autoCapitalize="none"
        selectionColor="#1688d8"
        underlineColorAndroid="transparent"
      />
    </View>
  );
}

export default function HomeScreen() {
  const [openSection, setOpenSection] = useState<SectionKey>("customer");

  const [customerName, setCustomerName] = useState("");
  const [customerStreet, setCustomerStreet] = useState("");
  const [city, setCity] = useState("");
  const [stateValue, setStateValue] = useState("FL");
  const [zip, setZip] = useState("");
  const [customerPhone, setCustomerPhone] = useState("");

  const [proposalDate, setProposalDate] = useState("");
  const [installStartDate, setInstallStartDate] = useState("");
  const [equipmentBrand, setEquipmentBrand] = useState("");
  const [permitCityCounty, setPermitCityCounty] = useState("");
  const [comfortConsultant, setComfortConsultant] = useState("");
  const [warehouseNumber, setWarehouseNumber] = useState("");

  const [systemName, setSystemName] = useState("");
  const [condenserModel, setCondenserModel] = useState("");
  const [airHandlerModel, setAirHandlerModel] = useState("");
  const [heatStrip, setHeatStrip] = useState("");
  const [extraEquipment1, setExtraEquipment1] = useState("");
  const [extraEquipment2, setExtraEquipment2] = useState("");

  const [thermostat, setThermostat] = useState("");
  const [uvLight, setUvLight] = useState("");
  const [dynamicAirCleaner, setDynamicAirCleaner] = useState("");
  const [mediaAirCleaner, setMediaAirCleaner] = useState("");
  const [pureAir, setPureAir] = useState("");
  const [otherAccessory, setOtherAccessory] = useState("");

  const [materials, setMaterials] = useState<MaterialItem[]>([
    { label: "Disconnect box", cells: ["K25", "L25", "M25"], values: ["", "", ""], customValues: ["", "", ""] },
    { label: "Re-Connect High voltage", cells: ["K26", "L26", "M26"], values: ["", "", ""], customValues: ["", "", ""] },
    { label: "Re-Connect Low voltage", cells: ["K27", "L27", "M27"], values: ["", "", ""], customValues: ["", "", ""] },
    { label: "New outdoor pad", cells: ["K28", "L28", "M28"], values: ["", "", ""], customValues: ["", "", ""] },
    { label: "Vibration Pad", cells: ["K29", "L29", "M29"], values: ["", "", ""], customValues: ["", "", ""] },
    { label: "Flush and Vacuum LineSet", cells: ["K30", "L30", "M30"], values: ["", "", ""], customValues: ["", "", ""] },
    { label: "Refrigerant filter dryer", cells: ["K31", "L31", "M31"], values: ["", "", ""], customValues: ["", "", ""] },
    { label: "Recover refrigerant", cells: ["K32", "L32", "M32"], values: ["", "", ""], customValues: ["", "", ""] },
    { label: "Mastic seal plenums", cells: ["K33", "L33", "M33"], values: ["", "", ""], customValues: ["", "", ""] },
    { label: "Modify supply plenum", cells: ["K34", "L34", "M34"], values: ["", "", ""], customValues: ["", "", ""] },
    { label: "Modify return plenum", cells: ["K35", "L35", "M35"], values: ["", "", ""], customValues: ["", "", ""] },
    { label: "Mastic and seal return/supply plenum up to 4 ft", cells: ["K36", "L36", "M36"], values: ["", "", ""], customValues: ["", "", ""] },
    { label: "Insulate ductwork", cells: ["R25", "S25", "T25"], values: ["", "", ""], customValues: ["", "", ""] },
    { label: "Old equipment removal", cells: ["R26", "S26", "T26"], values: ["", "", ""], customValues: ["", "", ""] },
    { label: "New supply ducts", cells: ["R27", "S27", "T27"], values: ["", "", ""], customValues: ["", "", ""] },
    { label: "New return ducts", cells: ["R28", "S28", "T28"], values: ["", "", ""], customValues: ["", "", ""] },
    { label: "New grills", cells: ["R29", "S29", "T29"], values: ["", "", ""], customValues: ["", "", ""] },
    { label: "Clean-Up", cells: ["R30", "S30", "T30"], values: ["", "", ""], customValues: ["", "", ""] },
    { label: "New condensate piping", cells: ["R31", "S31", "T31"], values: ["", "", ""], customValues: ["", "", ""] },
    { label: "Condensate safety switch", cells: ["R32", "S32", "T32"], values: ["", "", ""], customValues: ["", "", ""] },
    { label: "Condensate Pump w/ Sfty Switch", cells: ["R33", "S33", "T33"], values: ["", "", ""], customValues: ["", "", ""] },
    { label: "Drain pan with safety switch", cells: ["R34", "S34", "T34"], values: ["", "", ""], customValues: ["", "", ""] },
    { label: "1 inch filter rack", cells: ["R35", "S35", "T35"], values: ["", "", ""], customValues: ["", "", ""] },
    { label: "Other material", cells: ["R36", "S36", "T36"], values: ["", "", ""], customValues: ["", "", ""] },
  ]);

  const [annualServiceAgreement, setAnnualServiceAgreement] = useState("");
  const [additionalService1, setAdditionalService1] = useState("");
  const [additionalService2, setAdditionalService2] = useState("");

  const [scope1, setScope1] = useState("");
  const [scope2, setScope2] = useState("");
  const [scope3, setScope3] = useState("");
  const [scope4, setScope4] = useState("");
  const [scope5, setScope5] = useState("");

  const [maintenancePlan, setMaintenancePlan] = useState("");
  const [maintenanceNotes, setMaintenanceNotes] = useState("");
  const [coilCompressorWarranty, setCoilCompressorWarranty] = useState("");
  const [partsWarranty, setPartsWarranty] = useState("");
  const [laborWarranty, setLaborWarranty] = useState("");
  const [zoningWarranty, setZoningWarranty] = useState("");
  const [extendedWarranty, setExtendedWarranty] = useState("");
  const [otherWarranty, setOtherWarranty] = useState("");

  const [salePrice, setSalePrice] = useState("22200");
  const [coolBearDiscount, setCoolBearDiscount] = useState("1200");
  const [managerDiscount, setManagerDiscount] = useState("3100");
  const [paymentTerms, setPaymentTerms] = useState("");
  const [nameOnCard, setNameOnCard] = useState("");
  const [accountNumber, setAccountNumber] = useState("");
  const [expDate, setExpDate] = useState("");
  const [checkNumber, setCheckNumber] = useState("");
  const [memberNumber, setMemberNumber] = useState("");
  const [paymentNotes, setPaymentNotes] = useState("");

  const cleanNumber = (value: string) => {
    return Number(String(value).replace(/[^0-9.-]/g, "") || 0);
  };

  const formatDateInput = (text: string) => {
    const digits = text.replace(/\D/g, "").slice(0, 8);

    if (digits.length <= 2) {
      return digits;
    }

    if (digits.length <= 4) {
      return `${digits.slice(0, 2)}/${digits.slice(2)}`;
    }

    return `${digits.slice(0, 2)}/${digits.slice(2, 4)}/${digits.slice(4)}`;
  };

  const money = (value: string | number) => {
    const number = Number(value || 0);
    return number.toLocaleString("en-US", {
      style: "currency",
      currency: "USD",
    });
  };

  const totalInvestment =
    cleanNumber(salePrice) -
    cleanNumber(coolBearDiscount) -
    cleanNumber(managerDiscount);

  const shopCard = totalInvestment * 0.1;
  const executiveMembership = totalInvestment * 0.02;
  const costcoVisa = totalInvestment * 0.02;
  const netInvestment =
    totalInvestment - shopCard - executiveMembership - costcoVisa;

  const writeCell = (
    sheet: XLSX.WorkSheet,
    cell: string,
    value: string | number,
    type: "s" | "n" = "s"
  ) => {
    if (value === "" || value === null || value === undefined) return;

    sheet[cell] = {
      t: type,
      v: value,
    };
  };

  const setMaterialValue = (
    materialIndex: number,
    columnIndex: number,
    value: GridValue
  ) => {
    setMaterials((current) =>
      current.map((item, index) => {
        if (index !== materialIndex) return item;

        const nextValues = [...item.values];
        nextValues[columnIndex] = value;

        return {
          ...item,
          values: nextValues,
        };
      })
    );
  };

  const setMaterialCustomValue = (
    materialIndex: number,
    columnIndex: number,
    value: string
  ) => {
    setMaterials((current) =>
      current.map((item, index) => {
        if (index !== materialIndex) return item;

        const nextCustomValues = [...item.customValues];
        nextCustomValues[columnIndex] = value;

        return {
          ...item,
          customValues: nextCustomValues,
        };
      })
    );
  };

  const getMaterialOutput = (item: MaterialItem, columnIndex: number) => {
    const value = item.values[columnIndex];

    if (value === "CUSTOM") {
      return item.customValues[columnIndex];
    }

    return value;
  };

  const handleGenerate = async () => {
    try {
      console.log("Generate button pressed");

      if (Platform.OS === "web") {
        window.alert(
          "Button works. Excel generation needs iPad/iPhone through Expo Go."
        );
        return;
      }

      const templateAsset = Asset.fromModule(
        require("../../assets/templates/CostcoProposalTemplate.xlsx")
      );

      await templateAsset.downloadAsync();

      if (!templateAsset.localUri) {
        throw new Error("Could not load Excel template.");
      }

      const templateBase64 = await FileSystem.readAsStringAsync(
        templateAsset.localUri,
        {
          encoding: "base64" as any,
        }
      );

      const workbook = XLSX.read(templateBase64, { type: "base64" });
      const firstSheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[firstSheetName];

      writeCell(sheet, "B10", customerName);
      writeCell(sheet, "B11", customerStreet);
      writeCell(sheet, "B12", city);
      writeCell(sheet, "F12", stateValue);
      writeCell(sheet, "H12", zip);
      writeCell(sheet, "B13", customerPhone);

      writeCell(sheet, "P10", proposalDate);
      writeCell(sheet, "P11", installStartDate);
      writeCell(sheet, "P12", equipmentBrand);
      writeCell(sheet, "P13", permitCityCounty);
      writeCell(sheet, "P14", comfortConsultant);
      writeCell(sheet, "P15", warehouseNumber);

      writeCell(sheet, "K18", systemName);
      writeCell(sheet, "C19", condenserModel);
      writeCell(sheet, "C20", airHandlerModel);
      writeCell(sheet, "C21", heatStrip);
      writeCell(sheet, "C22", extraEquipment1);
      writeCell(sheet, "C23", extraEquipment2);

      writeCell(sheet, "C25", thermostat);
      writeCell(sheet, "C26", uvLight);
      writeCell(sheet, "C27", dynamicAirCleaner);
      writeCell(sheet, "C28", mediaAirCleaner);
      writeCell(sheet, "C29", pureAir);
      writeCell(sheet, "C30", otherAccessory);

      materials.forEach((item) => {
        item.cells.forEach((cell, columnIndex) => {
          writeCell(sheet, cell, getMaterialOutput(item, columnIndex));
        });
      });

      writeCell(sheet, "C34", annualServiceAgreement);
      writeCell(sheet, "C35", additionalService1);
      writeCell(sheet, "C36", additionalService2);

      writeCell(sheet, "A39", scope1);
      writeCell(sheet, "A40", scope2);
      writeCell(sheet, "A41", scope3);
      writeCell(sheet, "A42", scope4);
      writeCell(sheet, "A43", scope5);

      writeCell(sheet, "A46", maintenancePlan);
      writeCell(sheet, "A47", maintenanceNotes);
      writeCell(sheet, "C49", coilCompressorWarranty);
      writeCell(sheet, "C50", partsWarranty);
      writeCell(sheet, "C51", laborWarranty);
      writeCell(sheet, "C52", zoningWarranty);
      writeCell(sheet, "C53", extendedWarranty);
      writeCell(sheet, "C54", otherWarranty);

      writeCell(sheet, "P54", cleanNumber(salePrice), "n");
      writeCell(sheet, "P55", cleanNumber(coolBearDiscount), "n");
      writeCell(sheet, "P56", cleanNumber(managerDiscount), "n");

      writeCell(sheet, "P58", totalInvestment, "n");
      writeCell(sheet, "J54", shopCard, "n");
      writeCell(sheet, "J55", executiveMembership, "n");
      writeCell(sheet, "J56", costcoVisa, "n");
      writeCell(sheet, "G59", netInvestment, "n");
      writeCell(sheet, "P60", totalInvestment, "n");

      writeCell(sheet, "C58", paymentTerms);
      writeCell(sheet, "C60", nameOnCard);
      writeCell(sheet, "C61", accountNumber);
      writeCell(sheet, "H61", expDate);
      writeCell(sheet, "C62", checkNumber);
      writeCell(sheet, "H62", memberNumber);
      writeCell(sheet, "C63", paymentNotes);

      const outputBase64 = XLSX.write(workbook, {
        type: "base64",
        bookType: "xlsx",
      });

      const safeCustomerName =
        customerName.trim().replace(/[^a-z0-9]/gi, "_") || "Customer";

      const outputPath =
        FileSystem.documentDirectory +
        `Cool_Bear_Proposal_${safeCustomerName}.xlsx`;

      await FileSystem.writeAsStringAsync(outputPath, outputBase64, {
        encoding: "base64" as any,
      });

      const canShare = await Sharing.isAvailableAsync();

      if (!canShare) {
        Alert.alert("Proposal Created", `Saved to: ${outputPath}`);
        return;
      }

      await Sharing.shareAsync(outputPath, {
        mimeType:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        dialogTitle: "Share Cool Bear Proposal",
        UTI: "org.openxmlformats.spreadsheetml.sheet",
      });
    } catch (error) {
      console.error(error);
      Alert.alert(
        "Error",
        error instanceof Error
          ? error.message
          : "Something went wrong generating the Excel proposal."
      );
    }
  };

  const SectionButton = ({
    id,
    title,
  }: {
    id: SectionKey;
    title: string;
  }) => (
    <TouchableOpacity
      style={[
        styles.sectionButton,
        openSection === id && styles.sectionButtonActive,
      ]}
      onPress={() => setOpenSection(id)}
    >
      <Text
        selectable={false}
        style={[
          styles.sectionButtonText,
          openSection === id && styles.sectionButtonTextActive,
        ]}
      >
        {title}
      </Text>
    </TouchableOpacity>
  );

  const renderCustomerSection = () => (
    <View style={styles.card}>
      <Text selectable={false} style={styles.sectionTitle}>Customer / Installation Info</Text>

      <Field label="Customer Name" value={customerName} onChangeText={setCustomerName} />
      <Field label="Customer Street" value={customerStreet} onChangeText={setCustomerStreet} />
      <Field label="City" value={city} onChangeText={setCity} />

      <View style={styles.row}>
        <View style={styles.rowItemLeft}>
          <Field label="State" value={stateValue} onChangeText={setStateValue} />
        </View>
        <View style={styles.rowItemRight}>
          <Field label="Zip" value={zip} onChangeText={setZip} keyboardType="numeric" />
        </View>
      </View>

      <Field label="Customer Phone" value={customerPhone} onChangeText={setCustomerPhone} keyboardType="phone-pad" />

      <Field
        label="Date"
        value={proposalDate}
        onChangeText={(text: string) => setProposalDate(formatDateInput(text))}
        placeholder="MM/DD/YYYY"
        keyboardType="numeric"
      />

      <Field
        label="Installation Start Date"
        value={installStartDate}
        onChangeText={(text: string) => setInstallStartDate(formatDateInput(text))}
        placeholder="MM/DD/YYYY"
        keyboardType="numeric"
      />

      <Field label="Equipment Brand" value={equipmentBrand} onChangeText={setEquipmentBrand} />
      <Field label="Permit City/County" value={permitCityCounty} onChangeText={setPermitCityCounty} />
      <Field label="Comfort Consultant" value={comfortConsultant} onChangeText={setComfortConsultant} />
      <Field label="Warehouse Number" value={warehouseNumber} onChangeText={setWarehouseNumber} />
    </View>
  );

  const renderEquipmentSection = () => (
    <View style={styles.card}>
      <Text selectable={false} style={styles.sectionTitle}>Equipment to be Installed</Text>

      <Field label="System Name" value={systemName} onChangeText={setSystemName} />
      <Field label="Condenser Model" value={condenserModel} onChangeText={setCondenserModel} />
      <Field label="Air Handler Model" value={airHandlerModel} onChangeText={setAirHandlerModel} />
      <Field label="Heat Strip" value={heatStrip} onChangeText={setHeatStrip} />
      <Field label="Extra Equipment Line 1" value={extraEquipment1} onChangeText={setExtraEquipment1} />
      <Field label="Extra Equipment Line 2" value={extraEquipment2} onChangeText={setExtraEquipment2} />
    </View>
  );

  const renderAccessoriesSection = () => (
    <View style={styles.card}>
      <Text selectable={false} style={styles.sectionTitle}>Additional Accessories</Text>

      <Field label="Thermostat" value={thermostat} onChangeText={setThermostat} />
      <Field label="UV Light" value={uvLight} onChangeText={setUvLight} />
      <Field label="Dynamic Air Cleaner" value={dynamicAirCleaner} onChangeText={setDynamicAirCleaner} />
      <Field label={"5\" Media Air Cleaner"} value={mediaAirCleaner} onChangeText={setMediaAirCleaner} />
      <Field label="PureAir" value={pureAir} onChangeText={setPureAir} />
      <Field label="Other Accessory" value={otherAccessory} onChangeText={setOtherAccessory} />
    </View>
  );

  const renderMaterialsSection = () => (
    <View style={styles.card}>
      <Text selectable={false} style={styles.sectionTitle}>Additional Materials Checklist</Text>
      <Text selectable={false} style={styles.helperText}>
        Each row writes to System 1, System 2, and System 3 cells.
      </Text>

      {materials.map((item, materialIndex) => (
        <View key={item.label} style={styles.materialRow}>
          <Text selectable={false} style={styles.materialTitle}>{item.label}</Text>

          {[0, 1, 2].map((columnIndex) => (
            <View key={columnIndex} style={styles.materialSystemBlock}>
              <Text selectable={false} style={styles.systemLabel}>System {columnIndex + 1}</Text>

              <View style={styles.segmentRow}>
                {(["", "Y", "N", "CUSTOM"] as GridValue[]).map((option) => (
                  <TouchableOpacity
                    key={option || "blank"}
                    style={[
                      styles.segmentButton,
                      item.values[columnIndex] === option &&
                        styles.segmentButtonActive,
                    ]}
                    onPress={() =>
                      setMaterialValue(materialIndex, columnIndex, option)
                    }
                  >
                    <Text
                      selectable={false}
                      style={[
                        styles.segmentText,
                        item.values[columnIndex] === option &&
                          styles.segmentTextActive,
                      ]}
                    >
                      {option === ""
                        ? "Blank"
                        : option === "CUSTOM"
                        ? "Custom"
                        : option}
                    </Text>
                  </TouchableOpacity>
                ))}
              </View>

              {item.values[columnIndex] === "CUSTOM" && (
                <TextInput
                  style={fieldStyles.input}
                  value={item.customValues[columnIndex]}
                  onChangeText={(text: string) =>
                    setMaterialCustomValue(materialIndex, columnIndex, text)
                  }
                  placeholder="Custom value"
                  placeholderTextColor="#7f8b96"
                  autoCorrect={false}
                  spellCheck={false}
                  autoCapitalize="none"
                  selectionColor="#1688d8"
                  underlineColorAndroid="transparent"
                />
              )}
            </View>
          ))}
        </View>
      ))}
    </View>
  );

  const renderServicesSection = () => (
    <View style={styles.card}>
      <Text selectable={false} style={styles.sectionTitle}>Additional Services</Text>

      <Field label="Annual Service Agreement" value={annualServiceAgreement} onChangeText={setAnnualServiceAgreement} />
      <Field label="Additional Service Line 1" value={additionalService1} onChangeText={setAdditionalService1} />
      <Field label="Additional Service Line 2" value={additionalService2} onChangeText={setAdditionalService2} />
    </View>
  );

  const renderScopeSection = () => (
    <View style={styles.card}>
      <Text selectable={false} style={styles.sectionTitle}>Scope of Work</Text>

      <Field label="Scope Line 1" value={scope1} onChangeText={setScope1} multiline />
      <Field label="Scope Line 2" value={scope2} onChangeText={setScope2} multiline />
      <Field label="Scope Line 3" value={scope3} onChangeText={setScope3} multiline />
      <Field label="Scope Line 4" value={scope4} onChangeText={setScope4} multiline />
      <Field label="Scope Line 5" value={scope5} onChangeText={setScope5} multiline />
    </View>
  );

  const renderWarrantySection = () => (
    <View style={styles.card}>
      <Text selectable={false} style={styles.sectionTitle}>Warranty / Maintenance</Text>

      <Field label="Maintenance Plan" value={maintenancePlan} onChangeText={setMaintenancePlan} />
      <Field label="Maintenance Notes" value={maintenanceNotes} onChangeText={setMaintenanceNotes} multiline />
      <Field label="Coil / Compressor Warranty" value={coilCompressorWarranty} onChangeText={setCoilCompressorWarranty} />
      <Field label="Parts Warranty" value={partsWarranty} onChangeText={setPartsWarranty} />
      <Field label="Labor Warranty" value={laborWarranty} onChangeText={setLaborWarranty} />
      <Field label="Zoning Warranty" value={zoningWarranty} onChangeText={setZoningWarranty} />
      <Field label="Extended Warranty" value={extendedWarranty} onChangeText={setExtendedWarranty} />
      <Field label="Other Warranty" value={otherWarranty} onChangeText={setOtherWarranty} />
    </View>
  );

  const renderPaymentSection = () => (
    <View style={styles.card}>
      <Text selectable={false} style={styles.sectionTitle}>Payment / Pricing</Text>

      <Field label="Sale Price" value={salePrice} onChangeText={setSalePrice} keyboardType="numeric" />
      <Field label="Cool Bear Discount" value={coolBearDiscount} onChangeText={setCoolBearDiscount} keyboardType="numeric" />
      <Field label="Manager Discount" value={managerDiscount} onChangeText={setManagerDiscount} keyboardType="numeric" />

      <View style={styles.summaryBox}>
        <Text selectable={false} style={styles.summaryLine}>Total Investment: {money(totalInvestment)}</Text>
        <Text selectable={false} style={styles.summaryLine}>10% Shop Card: {money(shopCard)}</Text>
        <Text selectable={false} style={styles.summaryLine}>Executive Membership 2%: {money(executiveMembership)}</Text>
        <Text selectable={false} style={styles.summaryLine}>Costco Visa 2%: {money(costcoVisa)}</Text>
        <Text selectable={false} style={styles.summaryBig}>Net Out-of-Pocket: {money(netInvestment)}</Text>
      </View>

      <Field label="Payment Terms" value={paymentTerms} onChangeText={setPaymentTerms} multiline />
      <Field label="Name on Card" value={nameOnCard} onChangeText={setNameOnCard} />
      <Field label="Credit Card / Account #" value={accountNumber} onChangeText={setAccountNumber} />
      <Field label="Exp Date" value={expDate} onChangeText={setExpDate} />
      <Field label="Check #" value={checkNumber} onChangeText={setCheckNumber} />
      <Field label="Member #" value={memberNumber} onChangeText={setMemberNumber} />
      <Field label="Payment Notes" value={paymentNotes} onChangeText={setPaymentNotes} multiline />
    </View>
  );

  const renderOpenSection = () => {
    if (openSection === "customer") return renderCustomerSection();
    if (openSection === "equipment") return renderEquipmentSection();
    if (openSection === "accessories") return renderAccessoriesSection();
    if (openSection === "materials") return renderMaterialsSection();
    if (openSection === "services") return renderServicesSection();
    if (openSection === "scope") return renderScopeSection();
    if (openSection === "warranty") return renderWarrantySection();
    return renderPaymentSection();
  };

  return (
    <ScrollView style={styles.screen} contentContainerStyle={styles.container}>
      <View style={styles.header}>
        <Text selectable={false} style={styles.brand}>Cool Bear</Text>
        <Text selectable={false} style={styles.title}>Proposal Builder</Text>
        <Text selectable={false} style={styles.subtitle}>
          Section-by-section AC proposal editor
        </Text>
      </View>

      <View style={styles.navWrap}>
        <SectionButton id="customer" title="Customer" />
        <SectionButton id="equipment" title="Equipment" />
        <SectionButton id="accessories" title="Accessories" />
        <SectionButton id="materials" title="Materials" />
        <SectionButton id="services" title="Services" />
        <SectionButton id="scope" title="Scope" />
        <SectionButton id="warranty" title="Warranty" />
        <SectionButton id="payment" title="Payment" />
      </View>

      {renderOpenSection()}

      <TouchableOpacity style={styles.button} onPress={handleGenerate}>
        <Text selectable={false} style={styles.buttonText}>Generate Excel Proposal</Text>
      </TouchableOpacity>
    </ScrollView>
  );
}

const fieldStyles = StyleSheet.create({
  wrap: {
    width: "100%",
    marginBottom: 2,
    backgroundColor: "transparent",
  },
  label: {
    color: "#dce8f2",
    fontSize: 14,
    marginBottom: 6,
    marginTop: 10,
    fontWeight: "700",
    backgroundColor: "transparent",
    textShadowColor: "transparent",
    textDecorationLine: "none",
    overflow: "hidden",
    alignSelf: "flex-start",
  },
  input: {
    backgroundColor: "#071018",
    color: "#ffffff",
    borderWidth: 1,
    borderColor: "#2b4053",
    borderRadius: 12,
    paddingHorizontal: 14,
    paddingVertical: 12,
    fontSize: 16,
    textDecorationLine: "none",
  },
  textArea: {
    minHeight: 86,
  },
});

const styles = StyleSheet.create({
  screen: {
    flex: 1,
    backgroundColor: "#071018",
  },
  container: {
    padding: 18,
    paddingBottom: 50,
  },
  header: {
    marginBottom: 18,
    paddingTop: 18,
  },
  brand: {
    color: "#7ec8ff",
    fontSize: 22,
    fontWeight: "800",
    letterSpacing: 1,
    backgroundColor: "transparent",
  },
  title: {
    color: "#ffffff",
    fontSize: 34,
    fontWeight: "900",
    marginTop: 4,
    backgroundColor: "transparent",
  },
  subtitle: {
    color: "#a7b5c2",
    fontSize: 15,
    marginTop: 6,
    backgroundColor: "transparent",
  },
  navWrap: {
    flexDirection: "row",
    flexWrap: "wrap",
    marginBottom: 16,
  },
  sectionButton: {
    backgroundColor: "#111c26",
    borderWidth: 1,
    borderColor: "#1f3345",
    borderRadius: 999,
    paddingVertical: 10,
    paddingHorizontal: 13,
    marginRight: 8,
    marginBottom: 8,
  },
  sectionButtonActive: {
    backgroundColor: "#1688d8",
    borderColor: "#5bbdff",
  },
  sectionButtonText: {
    color: "#b9c7d3",
    fontWeight: "800",
    fontSize: 13,
    backgroundColor: "transparent",
  },
  sectionButtonTextActive: {
    color: "#ffffff",
  },
  card: {
    backgroundColor: "#111c26",
    borderRadius: 18,
    padding: 18,
    marginBottom: 18,
    borderWidth: 1,
    borderColor: "#1f3345",
  },
  sectionTitle: {
    color: "#ffffff",
    fontSize: 21,
    fontWeight: "900",
    marginBottom: 14,
    backgroundColor: "transparent",
    textShadowColor: "transparent",
  },
  helperText: {
    color: "#a7b5c2",
    fontSize: 14,
    lineHeight: 20,
    marginBottom: 14,
    backgroundColor: "transparent",
  },
  row: {
    flexDirection: "row",
  },
  rowItemLeft: {
    flex: 1,
    marginRight: 8,
  },
  rowItemRight: {
    flex: 1,
    marginLeft: 8,
  },
  materialRow: {
    backgroundColor: "#071018",
    borderRadius: 14,
    borderWidth: 1,
    borderColor: "#2b4053",
    padding: 14,
    marginBottom: 14,
  },
  materialTitle: {
    color: "#ffffff",
    fontSize: 16,
    fontWeight: "900",
    marginBottom: 10,
    backgroundColor: "transparent",
    textShadowColor: "transparent",
  },
  materialSystemBlock: {
    marginBottom: 12,
  },
  systemLabel: {
    color: "#b8dfff",
    fontSize: 13,
    fontWeight: "800",
    marginBottom: 6,
    backgroundColor: "transparent",
    textShadowColor: "transparent",
  },
  segmentRow: {
    flexDirection: "row",
    flexWrap: "wrap",
    marginBottom: 8,
  },
  segmentButton: {
    borderWidth: 1,
    borderColor: "#2b4053",
    borderRadius: 999,
    paddingVertical: 8,
    paddingHorizontal: 11,
    backgroundColor: "#111c26",
    marginRight: 6,
    marginBottom: 6,
  },
  segmentButtonActive: {
    backgroundColor: "#1688d8",
    borderColor: "#5bbdff",
  },
  segmentText: {
    color: "#cbd7e1",
    fontWeight: "800",
    fontSize: 12,
    backgroundColor: "transparent",
  },
  segmentTextActive: {
    color: "#ffffff",
  },
  summaryBox: {
    marginTop: 16,
    backgroundColor: "#071018",
    borderRadius: 12,
    padding: 14,
    borderWidth: 1,
    borderColor: "#2b4053",
  },
  summaryLine: {
    color: "#dce8f2",
    fontSize: 15,
    fontWeight: "700",
    marginBottom: 8,
    backgroundColor: "transparent",
  },
  summaryBig: {
    color: "#ffffff",
    fontSize: 19,
    fontWeight: "900",
    marginTop: 6,
    backgroundColor: "transparent",
  },
  button: {
    backgroundColor: "#1688d8",
    paddingVertical: 16,
    borderRadius: 16,
    alignItems: "center",
    marginTop: 4,
  },
  buttonText: {
    color: "#ffffff",
    fontSize: 17,
    fontWeight: "900",
    backgroundColor: "transparent",
  },
});
