using Microsoft.EntityFrameworkCore.Metadata;
using Microsoft.EntityFrameworkCore.Migrations;

namespace Project_V17.Data.Migrations
{
    public partial class ApplicationDBContext : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "FSApp",
                columns: table => new
                {
                    AppID = table.Column<int>(nullable: false)
                        .Annotation("SqlServer:ValueGenerationStrategy", SqlServerValueGenerationStrategy.IdentityColumn),
                    StaffFirstName = table.Column<string>(nullable: false),
                    StaffSurname = table.Column<string>(nullable: false),
                    Department = table.Column<string>(nullable: false),
                    Function = table.Column<string>(nullable: false),
                    CourseName = table.Column<string>(nullable: false),
                    Level = table.Column<int>(nullable: false),
                    Provider = table.Column<string>(nullable: false),
                    Details = table.Column<string>(nullable: false),
                    StartYear = table.Column<int>(nullable: false),
                    Mode = table.Column<string>(nullable: false),
                    Duration = table.Column<int>(nullable: false),
                    Cost = table.Column<float>(nullable: false),
                    Q1 = table.Column<string>(nullable: false),
                    Q2 = table.Column<string>(nullable: false),
                    Q3 = table.Column<string>(nullable: false),
                    Q4 = table.Column<string>(nullable: false),
                    Q5 = table.Column<string>(nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_FSApp", x => x.AppID);
                });
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "FSApp");
        }
    }
}
