/*  
 *  Author:         Rebecca Harris
 *  Date:           February 13, 2014
 *  Project:        Stoichiometry
 *  File name:      MainWindow.xaml.cs
 *  Description:    Provides the interface for the user to interact with to create
 *                  a formula and get the molecular weight of the fomrula. Interacts
 *                  with a database that contains the weight of each element and
 *                  a table of saved molecules. 
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using System.Data;
using System.Data.OleDb;

using Stoichiometry;

namespace Stoichiometry_RH
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // Variables
        public Molecule molecule;
        private double weight;
        private bool clearClicked;

        public MainWindow()
        {
            InitializeComponent();

            // Create molecule object to populate the list of molecules from the Molecules table in the database
            molecule = new Molecule();
            cbFormula.ItemsSource = molecule.FormulaList;
            clearClicked = false;
        }

        // User pressed the Normalize button
        private void btnNormalize_Click(object sender, RoutedEventArgs e)
        {
            // Create molecule object and set its formula to the text the user entered into the combo box
            molecule = new Molecule();
            molecule.Formula = cbFormula.Text;

            // Checks that the formula is valid
            if (molecule.Valid() == true)
            {
                weight = 0.0;

                // Normalizes the formula then sets the combo box text to be the normalized version
                molecule.Normalize();
                cbFormula.Text = molecule.Formula;

                // Get the molecular weight for the formula listed
                weight = molecule.MolecularWeight;
                txtMolWeight.Text = Convert.ToString(weight);
            }
            else
            {
                MessageBox.Show("Invalid formula entered, please enter a valid formula.","Formula Error",MessageBoxButton.OK,MessageBoxImage.Error);
            }

            // Empty molecule object
            molecule = null;
        }

        // User pressed the Calculate button
        private void btnCalculate_Click(object sender, RoutedEventArgs e)
        {
            // Create molecule object and set its formula to the text the user entered into the combo box
            molecule = new Molecule();
            molecule.Formula = cbFormula.Text;

            // Checks that the formula is valid
            if (molecule.Valid() == true)
            {
                // Get the molecular weight for the formula listed
                weight = molecule.MolecularWeight;
                txtMolWeight.Text = Convert.ToString(weight);
            }
            else
            {
                MessageBox.Show("Invalid formula entered, please enter a valid formula.", "Formula Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            // Empty molecule object
            molecule = null;
        }

        // Closes down the program
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Environment.Exit(0);
        }

        // User pressed the Save button
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            // Create molecule object and set its formula to the text the user entered into the combo box
            molecule = new Molecule();
            molecule.Formula = cbFormula.Text;

            // Checks that the formula is valid
            if (molecule.Valid() == true)
            {
                // Ask user if they want to permanently save the formula into the database
                MessageBoxResult result = MessageBox.Show("Would you like to permanently store this formula in the database?", "Save Formula", MessageBoxButton.YesNo, MessageBoxImage.Question);

                // Continues if they selected Yes, does nothing if they selected No
                if (result == MessageBoxResult.Yes)
                {
                    // Attempts to save the formula in the combo box into the database
                    if (molecule.Save() == true)
                    {
                        // Re-populate the combo box to include the new formula
                        cbFormula.ItemsSource = molecule.FormulaList;

                        // Lets the user know the formula saved successfully into the database
                        MessageBox.Show("Formula saved successfully.", "Formula Saved", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        // Lets the user know the formula did not save into the database
                        MessageBox.Show("Error: Formula was unable to save.", "Formula Not Saved", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            else
            {
                MessageBox.Show("Invalid formula entered, please enter a valid formula.", "Formula Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            // Empty molecule object
            molecule = null;
        }

        // User pressed the Clear button
        private void btnClear_Click(object sender, RoutedEventArgs e)
        {
            // Clears the current selection of the combo box and the molecular weight to be empty
            clearClicked = true;
            cbFormula.SelectedIndex = -1;
            cbFormula.Text = "";
            txtMolWeight.Text = "";
            clearClicked = false;
        }

        // User click on the combo box drop down button
        private void cbFormula_DropDownClosed(object sender, EventArgs e)
        {
            if (cbFormula.SelectedItem != null)
            {
                // Check that this event didn't trigger due to the Clear button being clicked
                if (clearClicked == false)
                {
                    // Create molecule object and set its formula to the text the user entered into the combo box
                    molecule = new Molecule();
                    molecule.Formula = cbFormula.SelectedItem.ToString();

                    // Checks that the formula is valid
                    if (molecule.Valid() == true)
                    {
                        // Get the molecular weight for the formula the user selected
                        weight = molecule.MolecularWeight;
                        txtMolWeight.Text = Convert.ToString(weight);
                    }
                }
            }
        }
    }
}
