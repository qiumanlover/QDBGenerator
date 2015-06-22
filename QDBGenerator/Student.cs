namespace QDBGenerator
{
    using System;

    internal class Student
    {
        private string grade;
        private string id;
        private string layer;
        private string name;
        private string origin;
        private string speciality;

        public Student()
        {
            this.id = string.Empty;
            this.name = string.Empty;
            this.origin = string.Empty;
            this.layer = string.Empty;
            this.speciality = string.Empty;
            this.grade = string.Empty;
        }

        public Student(string id, string name, string origin, string layer, string speciality, string grade)
        {
            this.id = string.Empty;
            this.name = string.Empty;
            this.origin = string.Empty;
            this.layer = string.Empty;
            this.speciality = string.Empty;
            this.grade = string.Empty;
            this.Id = id;
            this.Name = name;
            this.Origin = origin;
            this.Layer = layer;
            this.Speciality = speciality;
            this.Grade = grade;
        }

        public override string ToString()
        {
            return string.Format("{0} | {1} | {2} | {3} | {4} | {5}", new object[] { this.Origin, this.Grade, this.Layer, this.Speciality, this.Id, this.Name });
        }

        public string Grade
        {
            get
            {
                return this.grade;
            }
            set
            {
                this.grade = value;
            }
        }

        public string Id
        {
            get
            {
                return this.id;
            }
            set
            {
                this.id = value;
            }
        }

        public string Layer
        {
            get
            {
                return this.layer;
            }
            set
            {
                this.layer = value;
            }
        }

        public string Name
        {
            get
            {
                return this.name;
            }
            set
            {
                this.name = value;
            }
        }

        public string Origin
        {
            get
            {
                return this.origin;
            }
            set
            {
                this.origin = value;
            }
        }

        public string Speciality
        {
            get
            {
                return this.speciality;
            }
            set
            {
                this.speciality = value;
            }
        }
    }
}

