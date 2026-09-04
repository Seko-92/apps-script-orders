/**
 * fixture-getitem.js — one GetItem fixture, shared by both parser tests.
 *
 * NOT a copy of a real eBay payload — hand-built to carry the exact traps the
 * plan names, so it proves parser LOGIC. Element order follows eBay's schema
 * order (PictureDetails, ShippingDetails, then Variations near the end), which
 * is what makes the picture-leak case realistic rather than contrived.
 *
 * Lives in its own file because two copies of a fixture is the drift this
 * codebase keeps paying for.
 */

const XML_KIT = `<?xml version="1.0" encoding="utf-8"?>
<GetItemResponse xmlns="urn:ebay:apis:eBLBaseComponents">
<Ack>Success</Ack>
<Item>
  <ItemID>153889746781</ItemID>
  <SKU>157860</SKU>
  <Title>Engine Overhaul Kit</Title>
  <ConditionID>1000</ConditionID>
  <ConditionDisplayName>New</ConditionDisplayName>
  <Country>US</Country>
  <Location>Houston, Texas</Location>
  <PostalCode>77passed</PostalCode>
  <ListingType>FixedPriceItem</ListingType>
  <ListingDuration>GTC</ListingDuration>
  <StartPrice currencyID="USD">980.00</StartPrice>
  <Quantity>7</Quantity>
  <HideFromSearch>true</HideFromSearch>
  <ReasonHideFromSearch>OutOfStock</ReasonHideFromSearch>
  <OutOfStockControl>true</OutOfStockControl>
  <BestOfferEnabled>false</BestOfferEnabled>
  <WatchCount>47</WatchCount>
  <DispatchTimeMax>0</DispatchTimeMax>
  <ShipToLocations>US</ShipToLocations>
  <PictureDetails>
    <PictureURL>https://i.ebayimg.com/REAL-1.jpg</PictureURL>
    <PictureURL>https://i.ebayimg.com/REAL-2.jpg</PictureURL>
  </PictureDetails>
  <PrimaryCategory>
    <CategoryID>33615</CategoryID>
    <CategoryName>Heavy Equipment Parts</CategoryName>
  </PrimaryCategory>
  <Seller>
    <UserID>hqmotorservice</UserID>
    <FeedbackScore>4821</FeedbackScore>
    <PositiveFeedbackPercent>100.0</PositiveFeedbackPercent>
  </Seller>
  <SellingStatus>
    <CurrentPrice currencyID="USD">980.00</CurrentPrice>
    <QuantitySold>3</QuantitySold>
    <ListingStatus>Active</ListingStatus>
  </SellingStatus>
  <ShippingDetails>
    <ShippingType>Flat</ShippingType>
    <CalculatedShippingRate>
      <WeightMajor unit="lbs" measurementSystem="English">99</WeightMajor>
      <WeightMinor unit="oz" measurementSystem="English">0</WeightMinor>
    </CalculatedShippingRate>
    <ShippingServiceOptions>
      <ShippingService>ShippingMethodStandard</ShippingService>
      <ShippingServiceCost currencyID="USD">0.0</ShippingServiceCost>
      <FreeShipping>true</FreeShipping>
    </ShippingServiceOptions>
    <ShippingPackageDetails>
      <PackageDepth measurementSystem="English" unit="inches">10.00</PackageDepth>
      <PackageLength measurementSystem="English" unit="inches">20.00</PackageLength>
      <PackageWidth measurementSystem="English" unit="inches">18.00</PackageWidth>
      <ShippingIrregular>false</ShippingIrregular>
      <ShippingPackage>None</ShippingPackage>
      <WeightMajor unit="lbs" measurementSystem="English">25</WeightMajor>
      <WeightMinor unit="oz" measurementSystem="English">0</WeightMinor>
    </ShippingPackageDetails>
  </ShippingDetails>
  <ListingDetails>
    <StartTime>2026-01-04T18:00:00.000Z</StartTime>
    <EndTime>2026-12-04T18:00:00.000Z</EndTime>
    <ViewItemURL>https://www.ebay.com/itm/153889746781</ViewItemURL>
  </ListingDetails>
  <ReturnPolicy>
    <ReturnsAccepted>ReturnsAccepted</ReturnsAccepted>
    <ReturnsWithin>30 Days</ReturnsWithin>
  </ReturnPolicy>
  <ItemSpecifics>
    <NameValueList>
      <Name>Compatible Equipment Type</Name>
      <Value>Ditch Witch</Value>
      <Value>Crawler Tractor</Value>
      <Value>JLG</Value>
      <Value>Genie</Value>
      <Value>Dynapac</Value>
      <Value>Boom Lift</Value>
      <Value>Crawler Dozer</Value>
    </NameValueList>
    <NameValueList>
      <Name>Model Year</Name>
      <Value>K-55</Value>
    </NameValueList>
    <NameValueList>
      <Name>Brand</Name>
      <Value>HQ</Value>
    </NameValueList>
  </ItemSpecifics>
  <Variations>
    <Pictures>
      <VariationSpecificPictureSet>
        <PictureURL>https://i.ebayimg.com/VARIATION-A.jpg</PictureURL>
        <PictureURL>https://i.ebayimg.com/VARIATION-B.jpg</PictureURL>
        <PictureURL>https://i.ebayimg.com/VARIATION-C.jpg</PictureURL>
      </VariationSpecificPictureSet>
    </Pictures>
    <VariationSpecificsSet>
      <NameValueList>
        <Name>Bore Size</Name>
        <Value>STD</Value>
      </NameValueList>
    </VariationSpecificsSet>
  </Variations>
</Item>
</GetItemResponse>`;

// Flat-rate listing: NO CalculatedShippingRate, NO ShippingPackageDetails.
// This is the Phase 2 negative test — parcel fields must come back EMPTY, not wrong.
const XML_FLAT = XML_KIT
  .replace(/<CalculatedShippingRate>[\s\S]*?<\/CalculatedShippingRate>/, '')
  .replace(/<ShippingPackageDetails>[\s\S]*?<\/ShippingPackageDetails>/, '');

const XML_FAIL = `<?xml version="1.0"?><GetItemResponse><Ack>Failure</Ack>
<Errors><ShortMessage>Call usage limit has been reached.</ShortMessage></Errors></GetItemResponse>`;


// Metric listing — the packageDimsUnit guard exists so this is VISIBLE as an
// anomaly instead of silently mixing cm into a column named "In".
const XML_METRIC = XML_KIT
  .replace(/unit="inches"/g, 'unit="centimeters"')
  .replace(/<WeightMajor unit="lbs"([^>]*)>25</, '<WeightMajor unit="kg"$1>11<');

module.exports = { XML_KIT, XML_FLAT, XML_METRIC, XML_FAIL };
