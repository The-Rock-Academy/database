class Invoice {
  constructor(ID, parentName, pupilName, email, lessons, trials, costPerLesson, instrumentHireCost, company, term, note = "", type) {
    this.parentName = parentName;
    this.pupilName = pupilName;
    this.email = email;
    this.lessons = lessons;
    this.costPerLesson = costPerLesson;
    this.instrumentHireCost = instrumentHireCost;
    if (["TRA", "TSA", "GML"].includes(company)) {
      this.company = company;
    } else {
      throw new Error("Cannot have billing company: " + company);
    }
    
    //This is how the invoice numbers will be determined
    this.number = ID;
    this.term = term;
    this.note = note;
    // This should be one of three  values term, holiday and band.
    this.type = type;
    this.trials = trials;
    this.updated = false; // Whether or not this invoice is an update of a previous invoice.
  }

  getCompanyInfo() {
    switch (this.company) {
      case 'TRA':
        return {
          name: "The Rock Academy",
          image: "https://static.wixstatic.com/media/23b2e9_562405ec4e3344949bd3a22468d76d6c~mv2.png/v1/fill/w_553,h_202,al_c,q_85,usm_0.66_1.00_0.01,enc_auto/ROCK%20ACADEMY%20-%20Black%20on%20transparent_tif.png",
          address: [["83 Ross Street, Kilbirnie"], ["Wellington, 6022"], ["021 565 750"]],
          bankAccount: "01-0504-0147022-00"
        }
        break;
      case 'TSA':
        return {
          name: "The Singing Academy",
          image: "https://static.wixstatic.com/media/23b2e9_562405ec4e3344949bd3a22468d76d6c~mv2.png/v1/fill/w_553,h_202,al_c,q_85,usm_0.66_1.00_0.01,enc_auto/ROCK%20ACADEMY%20-%20Black%20on%20transparent_tif.png",
          address: [["83 Ross Street, Kilbirnie"], ["Wellington, 6022"], ["021 565 750"]],
          bankAccount: "06-0574-0245285-00"
        }
        break;
      case 'GML':
        return {
          name: "Geoffs Music Lessons",
          image: "https://static.wixstatic.com/media/23b2e9_562405ec4e3344949bd3a22468d76d6c~mv2.png/v1/fill/w_553,h_202,al_c,q_85,usm_0.66_1.00_0.01,enc_auto/ROCK%20ACADEMY%20-%20Black%20on%20transparent_tif.png",
          address: [["83 Ross Street, Kilbirnie"], ["Wellington, 6022"], ["021 565 750"]],
          bankAccount: "06-0574-0800509-00"
        }
        break;
      default:
        throw("There is not a billing company for this invoice. Only managed to find '" + this.company + "' in the Database");
    }
  }

/**
 * This function will take the number of lessons and the name of pupils to produce a item object.
 * This object will have desc quantity and price.
 * It shall return an array of these objects
 */
  getCosts() {
    let costs = [
      {
        desc: (this.type == "term" ? "Music Lessons" : this.type == "shp" ? "School Holiday Programme" : "Band School for " + this.term),
        quantity: this.lessons,
        price: this.costPerLesson
      }
    ]

    if (this.instrumentHireCost != "") {
      costs.push({
        desc: ("Instrument hire"),
        quantity: this.lessons + this.trials,
        price: this.instrumentHireCost
      })
    }

    return costs;
  }

  // /**
  //  * Takes the first letter of a initial and turns it into a full string
  //  */
  // getInstrumentName() {
  //   switch(this.instrumentHire) {
  //     case 'P':
  //       return "Piano";
  //     case 'S':
  //       return "Singing";
  //     case 'B':
  //       return "Bass";
  //     case 'U':
  //       return "Ukulele";
  //     case 'G':
  //       return "Guitar";
  //   }
  // }
  // /**
  //  * Returns how much the instrument hireage will be costing.
  //  */
  // getInstrumentCost() {
  //   if (this.instrumentHire != "") {
  //     return this.instrumentHire.split(" ").length * 12.50
  //   } else {
  //     return 0
  //   }
  // }
}

function newInvoice(invoiceFolderID, parentName, pupilName, email, lessons, trials, costPerLesson, instrumentHire, company, term, type) {
  let ID = (new InvoiceFolder(invoiceFolderID)).getNumberOfInvoices() + 1;
  return new Invoice(ID, parentName, pupilName, email, lessons, trials, costPerLesson, instrumentHire, company, term, "", type)
}
