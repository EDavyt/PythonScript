locationPopUpHandler: (data: IRuleObjectNode) => {

        const profLiabCov: boolean = documentHelper.getValue( data,"['P360 Coverage And Premium Proteus']['P360 Additional Coverages Proteus']['BooleanIsProfessionalLiabilityCoverageSelected']") === 'Yes'
        const stopGapCov: boolean = documentHelper.getValue( data,"['P360 Coverage And Premium Proteus']['P360 Additional Coverages Proteus']['BooleanIsEmployersLiabilityStopGapSelected']") === 'Yes'
        const AICov: boolean = documentHelper.getValue( data,"['P360 Coverage And Premium Proteus']['P360 Additional Coverages Proteus']['BooleanIsAdditionalInsuredSelected']") === 'Yes'
        const blanketAICov: boolean = documentHelper.getValue( data,"['P360 Coverage And Premium Proteus']['P360 Additional Coverages Proteus']['BooleanIsBlanketAdditionalInsuredCoverageSelected']") === 'Yes'
        const earthMovementCov:boolean = documentHelper.getValue( data,"['P360 Coverage And Premium Proteus']['P360 Additional Coverages Proteus']['BooleanIsEarthMovementSelected']") ==='Yes'
        const policyState: string = documentHelper.getValue( data,"['P360 General Information']['P360 Policy Details']['PolicyState']"
        )
        
        const formchecker = (formstocheck: string[]) => {
            // here goes some logic
        }

        const locationArrayNode = documentHelper.getNodeArray( data, "['P360 Coverage And Premium Proteus']['P360 Location And Class Information']" ).value as IRuleObjectNode[];

        for (let location of locationArrayNode) {
            const locationpath = documentHelper.getNode(data, location.path);
            const postalCodeAddress = documentHelper.getValue(locationpath, "PostalCodeAddress");
            const locationState = documentHelper.getValue(locationpath, "USStateAbbreviationAddressState");

            // Brownyard forms (any of 85000, 85001, 85020, 85500, 85501)
            // -> Brownyard classes drive these forms visibility
            if (["85000", "85001", "85020", "85500", "85501"].includes(postalCodeAddress)) {
                //Brownyard class present -> visible, else Hidden
                const brownyardForms = [ 'CG 00 68', 'CG 21 41', 'CG 21 50', 'CG 21 51', 'CG 22 31', 'CG 24 10', 'CG 25 46', 'HBCS-1301', 'HBCS-1304', 'HBCS-2000', 'HBCS-3001', 'HBCS-3005', 'HBCS-3006',]
                formchecker(brownyardForms)
            }

            // Farm class (31005)
            if (postalCodeAddress === "31005") {
                const farmForms = ['HUD-FM 2002']
                formchecker(farmForms)
            }

            // Amusement classes (87000–87007)
            if (["87000", "87001", "87002", "87003", "87004", "87005", "87006", "87007"].includes(postalCodeAddress)) {
                // Class-driven amusement forms
                const amusementClassForms = [ 'HUD-GL 3079', 'HUD-GL 3084', 'HUD-GL 3086'];
                formchecker(amusementClassForms);

                // Amusement + AI / Blanket AI coverage
                if ((AICov || blanketAICov) && ["87000","87001","87002","87003","87004","87005","87006","87007"].includes(postalCodeAddress)) {
                    const amusementAIForms = ['HUD-GL 2039',  'HUD-GL 2040']
                    formchecker(amusementAIForms);
                }
                //when postal code is 87006
                if (postalCodeAddress === "87006") {
                    const amusement87006Forms = ['HUD-GL 3080'];
                    formchecker(amusement87006Forms);
                }
            }

            // Prof Liab NOT selected + special classes (18200, 44311, 44315)
            if (!profLiabCov && ["18200", "44311", "44315"].includes(postalCodeAddress)) {
                const profLiabClassForms = ['CG 22 90'];
                formchecker(profLiabClassForms);
            }

            //Location State
            //TODO

            //Liquor Check
            let classArray = documentHelper.getValue(location, "['P360 Class Proteus']")
            
            if(documentHelper.getValue(data,"['isLiquorSelected']")== "Yes"){
                for (const classItem of classArray) {
                    const liquorSelected:boolean = documentHelper.getValue(classItem, "['BooleanIsLiquorLiabilityCoverageSelected']") === 'Yes'
                
                    //This is gonna be complicated, because of the conditions TODO
                    //  'HUD-LL 1000', // Liquor Liability selected on any class (hasLiquor(data) === true) -> Mandatory, else Hidden
                    //  'CG 00 33',    // Liquor Liability selected on any class (hasLiquor(data) === true) -> Mandatory, else Hidden
                    //  'CG 28 57',    // Liquor selected AND (PolicyState === 'CT' OR any LocationState === 'CT') -> Mandatory;

                }
            }
        }

        // Stop Gap by state (coverage-driven, no class-code dependency)
        if (stopGapCov && policyState === 'ND') {
            const ndStopGapForms = ['CG 04 40'];
            formchecker(ndStopGapForms);
        }

        if (stopGapCov && policyState === 'OH') {
            const ohStopGapForms = ['CG 04 41'];
            formchecker(ohStopGapForms);
        }

        if (stopGapCov && policyState === 'WA') {
            const waStopGapForms = ['CG 04 42'];
            formchecker(waStopGapForms);
        }

        if (stopGapCov && policyState === 'WY') {
            const wyStopGapForms = ['CG 04 44'];
            formchecker(wyStopGapForms);
        }

        //Limited Earth Movement coverage  Not selected
        if (!earthMovementCov) {
            const earthMovementForms = ['HUD-GL 3081'];
            formchecker(earthMovementForms);
        }
    },