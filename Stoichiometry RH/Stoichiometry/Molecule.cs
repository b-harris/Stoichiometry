/*  
 *  Author:         Rebecca Harris
 *  Date:           February 13, 2014
 *  Project:        Stoichiometry
 *  File name:      Molecule.cs
 *  Description:    Holds all the properties and methods to connect to the Stoichiometry database,
 *                  validate the formula of the molecule, get the weight of the molecule, save the 
 *                  molecule to the database, to normalize the formula entered by the user, and
 *                  to populate the combo box with molecules listed in the Molecules table in the
 *                  database.
 */


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Data;
using System.Data.OleDb;

namespace Stoichiometry
{
    public class Molecule : IMolecule
    {
        // Database member variables 
        private DataSet _ds;
        private OleDbDataAdapter _adElements, _adMolecules;

        // Member variables
        private bool validFormula = false;
        private bool savedSuccessfully = false;
        private double weight = 0.0;
        private string formula;
        private List<string> formulaL = new List<string>();
        private string validElem = "";
        Dictionary<string, int> dicElems = new Dictionary<string, int>();

        // Initialize database connection on creation of Molecule object
        public Molecule()
        {
            initializeDBConnection();
        }

        // Clearing dictionary and DB variables (in case)
        ~Molecule() 
        {
            _ds = null;
            _adElements = null;
            _adMolecules = null;
            dicElems.Clear();
        }

        /// <summary>
        /// Returns and sets a string of the formula for the molecule.
        /// </summary>
        public string Formula
        {
            get { return formula; }
            set
            {
                if (formula != value)
                {
                    formula = value;
                }
            }
        }

        /// <summary>
        /// Returns a double of the weight of the molecule.
        /// </summary>
        public double MolecularWeight
        {
            get { return weight; }  
        }

        /// <summary>
        /// Returns the list of formulas found in the database as a string array.
        /// </summary>
        public string[] FormulaList
        {
            get { return formulaL.ToArray(); }
        }

        /// <summary>
        /// Gets the total weight of the molecule based on the weight of each element times the count of each element.
        /// </summary>
        private void getMolecularWeight()
        {
            double elemWeight = 0.0;

            try
            {
                foreach (KeyValuePair<string, int> i in dicElems)
                {
                    foreach (DataRow row in _ds.Tables["Elements"].Rows)
                    {
                        // If the symbol exists in the database
                        if (i.Key == row.Field<string>("Symbol"))
                        {
                            // If there is only 1 occurance of the element (used 0 to represent if they don't specify a #, i.e. H vs. H1), add it to elemWeight
                            if (i.Value == 0)
                                elemWeight = row.Field<double>("AtomicWeight");
                            else // If there's more than 1, multiply the weight of that element by the count of that element
                                elemWeight = i.Value * row.Field<double>("AtomicWeight");

                            weight += elemWeight;
                            break;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception thrown in Molecule getMolecularWeight method: " + e.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Checks if the formula entered is valid.
        /// </summary>
        /// <returns>True if formula is valid, false if formula is invalid.</returns>
        public bool Valid()
        {
            string tempForm = formula, tempElem = "", tempElemNum = "";
            int cNum = 0, formLen = 0;

            try
            {
                // Store length of formula to match with later
                formLen = tempForm.Length;

                // Loop through each character in the formula
                foreach (char c in tempForm)
                {
                    cNum += 1;

                    //  Determine if it's a letter or digit and excluse white spaces
                    if (char.IsLetterOrDigit(c) && !char.IsWhiteSpace(c))
                    {
                        // If it's the first character, check that it's not a digit and that it's uppercase
                        if (cNum == 1)
                        {
                            if (!char.IsLetter(c) || !char.IsUpper(c))
                            {
                                validFormula = false;
                                break;
                            }
                        }

                        // Character is a digit
                        if (char.IsDigit(c))
                        {
                            // Checks if the digit is just a 0 (invalid)
                            if (c == '0' && tempElemNum == "")
                            {
                                validFormula = false;
                                break;
                            }

                            // New digit
                            if (tempElemNum == "")
                            {
                                // Check element prior to the digit
                                if (tempElem != "")
                                {
                                    // Check that the element prior to digit is valid
                                    validFormula = checkElementInDB(tempElem);

                                    tempElem = "";
                                    tempElemNum = "" + c;

                                    // Exits foreach loop if element does not exist in DB
                                    if (validFormula == false)
                                    {
                                        break;
                                    }
                                }
                                else
                                {
                                    validFormula = false;
                                    break;
                                }
                            }
                            else // Digit greater than 9, concatenate digit variable with next digit
                            {
                                tempElemNum = tempElemNum + c;
                            }
                        }
                        // Character is a letter
                        else if (char.IsLetter(c))
                        {
                            // Previous character was a digit, adds the prev element and its digit value to dictionary
                            if (validElem != "" && tempElemNum != "")
                            {
                                AddElemToDic(validElem, Convert.ToInt32(tempElemNum));
                                validElem = "";
                            }

                            tempElemNum = "";

                            // Current character is start of a new element, add previous one to dictionary
                            if (char.IsUpper(c) && tempElem != "")
                            {
                                // Check that the previous element is valid
                                validFormula = checkElementInDB(tempElem);

                                // Add valid element to dictionary
                                if (validElem != "")
                                {
                                    AddElemToDic(validElem, 0);
                                    validElem = "";
                                }

                                // Exits foreach loop if element does not exist in DB
                                if (validFormula == false)
                                {
                                    break;
                                }

                                // Store new element letter
                                tempElem = "" + c;
                            }
                            else
                            {
                                // Concatenate element variable with new character (i.e., Ni)
                                tempElem = tempElem + c;
                            }

                            // End of the formula, last character is a letter
                            if (cNum == formLen)
                            {
                                validFormula = checkElementInDB(tempElem);

                                if (validElem != "")
                                {
                                    AddElemToDic(validElem, 0);
                                    validElem = "";
                                } 
                                
                                break;
                            }
                        }
                    }
                    else
                    {
                        validFormula = false;
                        break;
                    }
                }

                // If digit was at the end of formula, add that element
                if (validElem != "" && tempElemNum != "")
                {
                    AddElemToDic(validElem, Convert.ToInt32(tempElemNum));
                }
            } 
            catch(Exception e)
            {
                MessageBox.Show("Exception thrown in Molecule Valid method: " + e.Message,"Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // Gets 
            getMolecularWeight();

            // Returns true if formula is valid, false if it's invalid
            return validFormula;
        }

        /// <summary>
        /// Adds the element and the count of how many times that element occurs into the element dictionary variable
        /// </summary>
        /// <param name="elem"></param>
        /// <param name="num"></param>
        private void AddElemToDic(string elem, int num)
        {
            bool added = false;

            try
            {
                // If there are already elements in the dictionary
                if (dicElems.Count > 0)
                {
                    // Checks if the element already exists in the dictionary
                    // If it does, add the count of the element to the existing value in the dictionary
                    foreach (KeyValuePair<string, int> i in dicElems)
                    {
                        if (i.Key == elem)
                        {
                            if (num == 0)
                                num = 1;

                            if (i.Value == 0)
                                num = 2;

                            int newNum = i.Value + num;
                            dicElems[i.Key] = newNum;
                            added = true;
                            break;
                        }
                    }

                    // Add the element to the dictionary
                    if (added == false)
                    {
                        dicElems.Add(elem, num);
                    }
                }
                else
                {
                    // Add the element to the dictionary
                    dicElems.Add(elem, num);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception thrown in Molecule AddElemToDic method: " + e.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Groups elements that are repeated through the element (i.e., NNiN becomes N2Ni)
        /// </summary>
        public void Normalize()
        {
            string newElem = "";

            try
            {
                foreach (KeyValuePair<string, int> i in dicElems)
                {
                    // If they didn't specify 1, then only concatenate the element's symbol (i.e., Ni)
                    if (i.Value == 0)
                    {
                        newElem = newElem + i.Key;
                    }
                    // If they did specify 1+, then concatenate the element's symbol and the # (i.e., Ni3)
                    else
                    {
                        newElem = newElem + i.Key + i.Value;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception thrown in Molecule Normalize method: " + e.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // Set the molecule's formula to the normalized version
            this.Formula = newElem;
        }

        /// <summary>
        /// Saves the formula in the combo box text into the database.
        /// </summary>
        /// <returns>True if formula was saved, false is not</returns>
        public bool Save()
        {
            // Make sure there is a formula entered to save
            if (this.Formula != "")
            {
                try
                {
                    // Add molecule to the dataset 
                    DataRow newMolecule = _ds.Tables["Molecules"].NewRow();

                    // Set molecule formula
                    newMolecule["Formula"] = this.Formula;

                    if (this.weight == 0)
                        getMolecularWeight();

                    // Set molecule weight
                    newMolecule["MolecularWeight"] = this.weight;

                    _ds.Tables["Molecules"].Rows.Add(newMolecule);

                    // Update the physical database with changes made to the Molecules
                    // table in the dataset
                    _adMolecules.Update(_ds, "Molecules");
                    savedSuccessfully = true;

                    // Add the new molecule to the formula list
                    populateFormulaList();
                }
                catch (Exception e)
                {
                    savedSuccessfully = false;
                }
            }

            // Returns true if formula was saved, false is not
            return savedSuccessfully;
        }

        /// <summary>
        /// Checks if the element passed in is valid (exists in the database's Elements table).
        /// </summary>
        /// <param name="elem"></param>
        /// <returns>True if the element is a valid element, false if not.</returns>
        private bool checkElementInDB(string elem)
        {
            bool isValid = false;

            foreach (DataRow row in _ds.Tables["Elements"].Rows)
            {
                // Checks that the element exists in the Elements table in the database
                if (elem == row.Field<string>("Symbol"))
                {
                    // If it exists, then it's valid
                    isValid = true;
                    validElem = elem;
                    break;
                }
            }

            // Returns true if the element is valid, false is not
            return isValid;
        }

        /// <summary>
        /// Creates a connection to the Stoichiometry database.
        /// </summary>
        public void initializeDBConnection()
        {
            try
            {
                // Create a connection object for the Stoichiometry database
                OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Stoichiometry.accdb");

                // Create data adapter for Elements table
                _adElements = new OleDbDataAdapter();
                _adElements.SelectCommand = new OleDbCommand("SELECT * FROM Elements", conn);

                // Create data adapter for Molecules table
                _adMolecules = new OleDbDataAdapter();
                _adMolecules.SelectCommand = new OleDbCommand("SELECT * FROM Molecules", conn);
                _adMolecules.InsertCommand = new OleDbCommand("INSERT INTO Molecules (Formula, MolecularWeight) VALUES(@Formula,@MolecularWeight)", conn);

                // Create a parameter variable for the insert command
                _adMolecules.InsertCommand.Parameters.Add("@Formula", OleDbType.VarChar, -1, "Formula");
                _adMolecules.InsertCommand.Parameters.Add("@MolecularWeight", OleDbType.Double, -1, "MolecularWeight");

                // Create and fill the data set and close the connection 
                _ds = new DataSet();
                conn.Open();
                _adElements.Fill(_ds, "Elements");
                _adMolecules.Fill(_ds, "Molecules");
                conn.Close();

                // Populate the formula list with all molecules in the Molecules table in the database
                populateFormulaList();
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception thrown in Molecule initializeElements method: " + e.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Populate the formula list combo box with molecules from the Molecules table in the database
        /// </summary>
        private void populateFormulaList()
        {
            // Clear any previous formulas
            formulaL.Clear();

            // Populate the Formulas list variable
            foreach (DataRow row in _ds.Tables["Molecules"].Rows)
            {
                formulaL.Add(row.Field<String>("Formula"));
            }
        }
    }

    /// <summary>
    /// Interface for Molecule
    /// </summary>
    public interface IMolecule
    {
        string Formula { get; set; }
        bool Valid();
        double MolecularWeight { get; }
        void Normalize();
        bool Save();
        string[] FormulaList { get; }
    }
}
