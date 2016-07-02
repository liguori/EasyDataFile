using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDataFile.Import.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ValidateAttribute : Attribute
    {
        public ValidateAttribute(int maxlLenght, bool required, ValidationType typeToValidate)
        {
            this.MaxLenght = maxlLenght;
            this.Required = required;
            this.TypeToValidate = typeToValidate;
        }

        public int MaxLenght { get; set; }

        public bool Required { get; set; }

        public ValidationType TypeToValidate { get; set; }

        public enum ValidationType
        {
            Integer,
            Float,
            Date,
            String
        }

        public void Validate(string name, string value)
        {
            //Check if value is required and can't be null or empty
            if (string.IsNullOrWhiteSpace(value) && this.Required)
                throw new Exception("Campo '" + name + "' è un campo obbligatorio ma non è impostato");
            //Check if max lenght of the value is correct
            if (!string.IsNullOrWhiteSpace(value) && value.Length > this.MaxLenght)
                throw new Exception("Campo '" + name + "' ha una lunghezza non valida, deve essere di " + this.MaxLenght.ToString() + " caratteri");
            //Check the data type format
            if (!string.IsNullOrEmpty(value))
            {
                switch (this.TypeToValidate)
                {
                    case ValidationType.Date:
                        bool valido = true;
                        DateTime dvalOut;
                        if (!value.Contains("/")){
                            try
                            {
                                dvalOut = DateTime.FromOADate(double.Parse(value));
                            }
                            catch (Exception)
                            {
                                valido = false;
                            }

                        } else if (!DateTime.TryParse(value, out dvalOut))
                        {
                            valido = false;
                        }
                        if (valido == false) throw new Exception("Campo '" + name + "' contiene un valore non valido per il tipo 'Data'");
                        break;
                    case ValidationType.Integer:
                        long lvalOut;
                        if (!long.TryParse(value, out lvalOut))
                            throw new Exception("Campo '" + name + "' contiene un valore non valido per il tipo 'Numerico intero'");
                        break;
                    case ValidationType.Float:
                        double fvalOut;
                        if (!double.TryParse(value, out fvalOut))
                            throw new Exception("Campo '" + name + "' contiene un valore non valido per il tipo 'Numerico con virgola'");
                        break;
                }
            }
        }

    }
}