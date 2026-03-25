using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace TaxPersonnelManagement.Migrations
{
    /// <inheritdoc />
    public partial class AddRewardInfo : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "Departments",
                columns: table => new
                {
                    Id = table.Column<int>(type: "INTEGER", nullable: false)
                        .Annotation("Sqlite:Autoincrement", true),
                    Name = table.Column<string>(type: "TEXT", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Departments", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "Personnel",
                columns: table => new
                {
                    Id = table.Column<int>(type: "INTEGER", nullable: false)
                        .Annotation("Sqlite:Autoincrement", true),
                    StaffId = table.Column<string>(type: "TEXT", nullable: true),
                    FullName = table.Column<string>(type: "TEXT", nullable: false),
                    DateOfBirth = table.Column<DateTime>(type: "TEXT", nullable: true),
                    Gender = table.Column<string>(type: "TEXT", nullable: true),
                    PhoneNumber = table.Column<string>(type: "TEXT", nullable: true),
                    IdentityCardNumber = table.Column<string>(type: "TEXT", nullable: true),
                    IdentityCardPlace = table.Column<string>(type: "TEXT", nullable: true),
                    SocialSecurityNumber = table.Column<string>(type: "TEXT", nullable: true),
                    Email = table.Column<string>(type: "TEXT", nullable: true),
                    BirthPlace = table.Column<string>(type: "TEXT", nullable: true),
                    Ethnicity = table.Column<string>(type: "TEXT", nullable: true),
                    Religion = table.Column<string>(type: "TEXT", nullable: true),
                    Department = table.Column<string>(type: "TEXT", nullable: true),
                    Position = table.Column<string>(type: "TEXT", nullable: true),
                    RankCode = table.Column<string>(type: "TEXT", nullable: true),
                    RankName = table.Column<string>(type: "TEXT", nullable: true),
                    TaxAuthorityStartDate = table.Column<DateTime>(type: "TEXT", nullable: true),
                    StartDate = table.Column<DateTime>(type: "TEXT", nullable: true),
                    Status = table.Column<string>(type: "TEXT", nullable: true),
                    PositionDecisionDate = table.Column<DateTime>(type: "TEXT", nullable: true),
                    PositionCalculationDate = table.Column<DateTime>(type: "TEXT", nullable: true),
                    PositionYear = table.Column<string>(type: "TEXT", nullable: true),
                    DetailedWorkHistory = table.Column<string>(type: "TEXT", nullable: true),
                    RetirementDate = table.Column<DateTime>(type: "TEXT", nullable: true),
                    TotalAnnualLeaveDays = table.Column<int>(type: "INTEGER", nullable: false),
                    PartyEntryDate = table.Column<DateTime>(type: "TEXT", nullable: true),
                    PartyOfficialDate = table.Column<DateTime>(type: "TEXT", nullable: true),
                    EducationLevel = table.Column<string>(type: "TEXT", nullable: true),
                    Major = table.Column<string>(type: "TEXT", nullable: true),
                    University = table.Column<string>(type: "TEXT", nullable: true),
                    StateManagementLevel = table.Column<string>(type: "TEXT", nullable: true),
                    PoliticalTheoryLevel = table.Column<string>(type: "TEXT", nullable: true),
                    ITSkillLevel = table.Column<string>(type: "TEXT", nullable: true),
                    LanguageSkillLevel = table.Column<string>(type: "TEXT", nullable: true),
                    AvatarBase64 = table.Column<string>(type: "TEXT", nullable: true),
                    CurrentSalaryStep = table.Column<string>(type: "TEXT", nullable: true),
                    CurrentSalaryCoefficient = table.Column<double>(type: "REAL", nullable: false),
                    ExceedFramePercent = table.Column<double>(type: "REAL", nullable: false),
                    PositionAllowance = table.Column<string>(type: "TEXT", nullable: true),
                    SalaryReservationDeadline = table.Column<DateTime>(type: "TEXT", nullable: true),
                    NextSalaryStepDate = table.Column<DateTime>(type: "TEXT", nullable: true),
                    SalaryIncreaseDelayType = table.Column<string>(type: "TEXT", nullable: true),
                    ExpectedSalaryIncreaseDate = table.Column<DateTime>(type: "TEXT", nullable: true),
                    SalaryHistoryLog = table.Column<string>(type: "TEXT", nullable: true),
                    EmulationTitles = table.Column<string>(type: "TEXT", nullable: true),
                    RewardForms = table.Column<string>(type: "TEXT", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Personnel", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "Positions",
                columns: table => new
                {
                    Id = table.Column<int>(type: "INTEGER", nullable: false)
                        .Annotation("Sqlite:Autoincrement", true),
                    Name = table.Column<string>(type: "TEXT", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Positions", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "Ranks",
                columns: table => new
                {
                    Id = table.Column<int>(type: "INTEGER", nullable: false)
                        .Annotation("Sqlite:Autoincrement", true),
                    Code = table.Column<string>(type: "TEXT", nullable: false),
                    Name = table.Column<string>(type: "TEXT", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Ranks", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "RankSalarySpecs",
                columns: table => new
                {
                    Id = table.Column<int>(type: "INTEGER", nullable: false)
                        .Annotation("Sqlite:Autoincrement", true),
                    RankCode = table.Column<string>(type: "TEXT", nullable: false),
                    SalaryStep = table.Column<string>(type: "TEXT", nullable: false),
                    Coefficient = table.Column<double>(type: "REAL", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_RankSalarySpecs", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "SalaryDelayReasons",
                columns: table => new
                {
                    Id = table.Column<int>(type: "INTEGER", nullable: false)
                        .Annotation("Sqlite:Autoincrement", true),
                    Name = table.Column<string>(type: "TEXT", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_SalaryDelayReasons", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "Users",
                columns: table => new
                {
                    Id = table.Column<int>(type: "INTEGER", nullable: false)
                        .Annotation("Sqlite:Autoincrement", true),
                    Username = table.Column<string>(type: "TEXT", nullable: false),
                    PasswordHash = table.Column<string>(type: "TEXT", nullable: false),
                    FullName = table.Column<string>(type: "TEXT", nullable: false),
                    Role = table.Column<int>(type: "INTEGER", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Users", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "LeaveHistories",
                columns: table => new
                {
                    Id = table.Column<int>(type: "INTEGER", nullable: false)
                        .Annotation("Sqlite:Autoincrement", true),
                    PersonnelId = table.Column<int>(type: "INTEGER", nullable: false),
                    LeaveType = table.Column<string>(type: "TEXT", nullable: false),
                    StartDate = table.Column<DateTime>(type: "TEXT", nullable: false),
                    EndDate = table.Column<DateTime>(type: "TEXT", nullable: false),
                    DurationDays = table.Column<double>(type: "REAL", nullable: false),
                    Reason = table.Column<string>(type: "TEXT", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_LeaveHistories", x => x.Id);
                    table.ForeignKey(
                        name: "FK_LeaveHistories_Personnel_PersonnelId",
                        column: x => x.PersonnelId,
                        principalTable: "Personnel",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateTable(
                name: "SalaryRecords",
                columns: table => new
                {
                    Id = table.Column<int>(type: "INTEGER", nullable: false)
                        .Annotation("Sqlite:Autoincrement", true),
                    PersonnelId = table.Column<int>(type: "INTEGER", nullable: false),
                    SalaryLevel = table.Column<string>(type: "TEXT", nullable: false),
                    Coefficient = table.Column<decimal>(type: "TEXT", nullable: false),
                    DecisionDate = table.Column<DateTime>(type: "TEXT", nullable: false),
                    EffectiveDate = table.Column<DateTime>(type: "TEXT", nullable: false),
                    NextIncreaseDate = table.Column<DateTime>(type: "TEXT", nullable: false),
                    Note = table.Column<string>(type: "TEXT", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_SalaryRecords", x => x.Id);
                    table.ForeignKey(
                        name: "FK_SalaryRecords_Personnel_PersonnelId",
                        column: x => x.PersonnelId,
                        principalTable: "Personnel",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.InsertData(
                table: "Users",
                columns: new[] { "Id", "FullName", "PasswordHash", "Role", "Username" },
                values: new object[] { 1, "Administrator", "admin", 0, "admin" });

            migrationBuilder.CreateIndex(
                name: "IX_LeaveHistories_PersonnelId",
                table: "LeaveHistories",
                column: "PersonnelId");

            migrationBuilder.CreateIndex(
                name: "IX_SalaryRecords_PersonnelId",
                table: "SalaryRecords",
                column: "PersonnelId");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "Departments");

            migrationBuilder.DropTable(
                name: "LeaveHistories");

            migrationBuilder.DropTable(
                name: "Positions");

            migrationBuilder.DropTable(
                name: "Ranks");

            migrationBuilder.DropTable(
                name: "RankSalarySpecs");

            migrationBuilder.DropTable(
                name: "SalaryDelayReasons");

            migrationBuilder.DropTable(
                name: "SalaryRecords");

            migrationBuilder.DropTable(
                name: "Users");

            migrationBuilder.DropTable(
                name: "Personnel");
        }
    }
}
