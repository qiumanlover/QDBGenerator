namespace QDBGenerator
{
    using System;

    internal class Course
    {
        private string courseName;
        private string property;
        private string speciality;
        private string time;

        public Course()
        {
            this.speciality = string.Empty;
            this.courseName = string.Empty;
            this.property = string.Empty;
            this.time = string.Empty;
        }

        public Course(string speciality, string name, string property, string time)
        {
            this.speciality = string.Empty;
            this.courseName = string.Empty;
            this.property = string.Empty;
            this.time = string.Empty;
            this.Speciality = speciality;
            this.CourseName = name;
            this.Property = property;
            this.Time = time;
        }

        public override string ToString()
        {
            return string.Format("{0} | {1} | {2} | {3}", new object[] { this.Speciality, this.CourseName, this.Property, this.Time });
        }

        public string CourseName
        {
            get
            {
                return this.courseName;
            }
            set
            {
                this.courseName = value;
            }
        }

        public string Property
        {
            get
            {
                return this.property;
            }
            set
            {
                this.property = value;
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

        public string Time
        {
            get
            {
                return this.time;
            }
            set
            {
                this.time = value;
            }
        }
    }
}

