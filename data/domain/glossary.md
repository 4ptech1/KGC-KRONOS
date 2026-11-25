# The KRONOS glossary: welding, pressure-vessel, estimation and shop words (D133, E1)

In-house (D25): written by Claude through our pipeline from the words KGC's own documents use (the survey,
`docs/kgc-documents.md`), for Boss to read and correct. **It is for understanding requests, never for answers:** KRONOS
uses it to match the words a person types to the recipes it already has (the recipe finder, `compiler/kronos_finder.py`)
and, from March 2027, to teach the model KGC's way of writing requests. No number, rule or code value here is used in a
calculation; those come only from Boss's confirmed rules (`docs/engineering-knowledge-plan.md`).

How to read a row: **Term**, the word KRONOS files it under; **Means**, a plain meaning; **Also written as**, other ways
people write it, separated by `;` (single words are matched by the recipe finder, phrases help people and the model);
**Area**: welding, vessel, NDE, quality, estimation or shop.

| Term | Means | Also written as | Area |
| --- | --- | --- | --- |
| WPS | Welding Procedure Specification: the written instructions for making one kind of weld | weld procedure; welding procedure specification | welding |
| PQR | Procedure Qualification Record: the test weld and its results that qualify a WPS | WPQR; procedure qualification record; procedure test | welding |
| WPQ | Welder Performance Qualification: proof that one welder can weld to a WPS | welder qualification test; WQT; welder test | welding |
| process | the welding process, by name or ISO number | GTAW; TIG; 141; SMAW; MMA; 111; SAW; 121; GMAW; MIG; MAG; FCAW | welding |
| run | one pass of weld metal laid along the joint | pass; bead; layer; root run; fill run; cap run | welding |
| root | the first run at the bottom of the joint | root pass; root run | welding |
| cap | the last runs on the surface of the weld | capping; cover pass; final run | welding |
| backgouge | removing the root from the back side before welding that side | back gouging; back-gouge; backchip; back chipping | welding |
| heatinput | the energy put into the weld per length of weld, in kJ/mm | heat input; HI; arc energy | welding |
| current | the welding current, in amperes | amps; amperage; A | welding |
| voltage | the arc voltage, in volts | volts; V | welding |
| travelspeed | how fast the torch or electrode moves along the joint, in mm/min | travel speed; speed of travel | welding |
| polarity | the type of current and which side is positive | DCEN; DCEP; AC; current type | welding |
| position | the welding position | PA; PB; PC; PD; PE; PF; PG; 1G; 2G; 3G; 4G; 5G; 6G; flat; horizontal; vertical | welding |
| preheat | heating the joint before welding | preheating; preheat temperature | welding |
| interpass | the temperature of the joint between runs | interpass temperature | welding |
| pwht | post weld heat treatment: heating the finished weld to relieve stress | post weld heat treatment; stress relief; stress relieving; SR | welding |
| filler | the wire or rod that becomes weld metal | filler metal; filler wire; electrode; rod; consumable | welding |
| shielding | the gas that protects the weld pool | shielding gas; argon; Ar; backing gas; purge | welding |
| tungsten | the electrode of a TIG torch, which does not melt | tungsten electrode; EWTh; WT20 | welding |
| pno | the group a parent metal belongs to, for qualification | P-No; P number; material group; ISO 15608 group | welding |
| parentmetal | the metal being welded | parent material; base metal; BM | welding |
| thickness | the material thickness, in mm | thk; t; wall thickness | welding |
| qualifiedrange | the range a WPS covers, from what its PQR tested | qualified range; range qualified; validity | welding |
| haz | the heat affected zone beside the weld | heat affected zone; HAZ | welding |
| impact | the Charpy impact test of toughness | CVN; Charpy; impact test; notch toughness | welding |
| hardness | the hardness test of weld, HAZ and parent metal | HV10; HBW; HRC; BHN | welding |
| examiner | the body that witnesses and approves a WPS or PQR | examining body; third party; TPI; LR; Lloyd's | quality |
| shell | the cylinder of a pressure vessel | shell course; cylinder | vessel |
| head | the end closing a vessel | dished end; dish end; DE; 2:1 SE; ellipsoidal head; torispherical head | vessel |
| nozzle | a pipe branch on a vessel, with its flange | branch; connection; inlet; outlet; vent; drain | vessel |
| repad | a reinforcing pad around a nozzle | RF pad; reinforcement pad; pad | vessel |
| flange | the bolted ring at a nozzle or vessel end | WNRF; weld neck; blind; LWN; long weld neck | vessel |
| rating | the pressure class of a flange | class; 150#; 300#; 600#; pressure class | vessel |
| saddle | a support a horizontal vessel stands on | saddles; support | vessel |
| lug | a plate welded on for lifting | lifting lug; lugs | vessel |
| designpressure | the pressure a vessel is designed for | design pressure; DP; design pr | vessel |
| designtemp | the temperature a vessel is designed for | design temperature; DT; design temp | vessel |
| mawp | the maximum allowable working pressure of the finished vessel | MAWP | vessel |
| mdmt | the minimum design metal temperature | MDMT | vessel |
| ca | the corrosion allowance added to a thickness | corrosion allowance; CA | vessel |
| jointefficiency | the factor for how fully the welds are examined | joint efficiency; E | vessel |
| hydrotest | a pressure test with water | hydrostatic test; hydro; pressure test; HT | vessel |
| ustamp | the ASME code stamp on a vessel | U stamp; U-stamp; code stamp; ASME stamp | vessel |
| nace | sour (wet H2S) service rules for materials and welding | NACE; sour service; wet H2S; MR0175; MR0103 | vessel |
| hic | hydrogen induced cracking, and the test of plate against it | HIC; SSC | vessel |
| moc | the material of construction | MOC; material of construction; material | vessel |
| ttl | tangent line to tangent line: the shell length between the heads | TL to TL; TL-TL; T/T | vessel |
| od | the outside diameter | OD; outside diameter; outer diameter | vessel |
| id | the inside diameter | ID; inside diameter; bore | vessel |
| tag | the equipment tag number | tag no; tag number; equipment no | vessel |
| ga | the general arrangement drawing | GA; general arrangement; drawing; drg | vessel |
| rt | radiographic testing of welds | radiography; X-ray; RT-1 | NDE |
| ut | ultrasonic testing | ultrasonic; UT; TOFD; PAUT | NDE |
| mt | magnetic particle testing | MPI; MPT; magnetic particle | NDE |
| pt | dye penetrant testing | DPT; dye penetrant; LPT; penetrant | NDE |
| pmi | positive material identification | PMI | NDE |
| visual | visual examination of a weld | VT; visual inspection | NDE |
| qap | the quality assurance plan or inspection and test plan for a job | QAP; ITP; inspection and test plan | quality |
| mtc | the material test certificate from the mill | MTC; mill certificate; 3.1 certificate; MTR; test certificate | quality |
| mdr | the manufacturer's data report of an ASME vessel | MDR; U-1A | quality |
| databook | the final collection of a job's documents | data book; final data book; dossier | quality |
| enquiry | a customer's request for a price | inquiry; enquiry; RFQ; request for quotation | estimation |
| proposal | the technical offer sent to a customer | technical proposal; offer; quotation; quote | estimation |
| estimate | the internal working-out of cost and price | estimation; estimation sheet; costing | estimation |
| tq | a technical query or deviation sent with an offer | TQ; technical query; deviation | estimation |
| tdc | the technical delivery conditions of one item in a proposal | TDC; technical delivery conditions | estimation |
| bom | the bill of materials | BOM; bill of materials; BOM-RM | estimation |
| rawmaterial | material as bought, before cutting and machining | raw material; RM; raw weight; gross weight | estimation |
| finishedweight | the weight of the finished part | finished weight; fin wt; drawing weight; drg unit wt | estimation |
| conversion | the cost of making: labour, machining, NDT, PWHT, painting | conversion cost; fabrication cost; labour | estimation |
| overheads | the factory overheads added to cost | overhead; OH | estimation |
| contingency | an allowance for the unknown | contingency | estimation |
| margin | the mark-up on cost | margin; mark-up; markup; profit | estimation |
| exworks | the price at KGC's works, before freight | ex-works; EXW; ex works | estimation |
| gst | India's goods and services tax | GST; tax | estimation |
| inr | Indian rupees | INR; Rs; rupees | estimation |
| lstk | lump sum turnkey contractor | LSTK | shop |
| epc | the engineering, procurement and construction contractor, or the project management consultant | EPC; PMC; contractor | shop |
| enduser | the company that will use the equipment | end user; client; owner | shop |
| oem | the original maker of the equipment | OEM; original manufacturer | shop |
| tba | to be advised: not known yet | TBA; TBC; to be confirmed | shop |
| srno | the serial number column of a list | Sr.No.; Sr Nos; Sl.; S.No.; serial number | shop |
