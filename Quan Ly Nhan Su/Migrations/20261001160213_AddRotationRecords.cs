using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace TaxPersonnelManagement.Migrations
{
    /// <inheritdoc />
    public partial class AddRotationRecords : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<double>(
                name: "ExceedFramePercent",
                table: "SalaryRecords",
                type: "REAL",
                nullable: false,
                defaultValue: 0.0);

            migrationBuilder.AddColumn<string>(
                name: "RankCode",
                table: "SalaryRecords",
                type: "TEXT",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "SalaryStep",
                table: "SalaryRecords",
                type: "TEXT",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "DepartmentName",
                table: "Positions",
                type: "TEXT",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "GroupType",
                table: "Positions",
                type: "TEXT",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "CurrentResidence",
                table: "Personnel",
                type: "TEXT",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "Hometown",
                table: "Personnel",
                type: "TEXT",
                nullable: true);

            migrationBuilder.CreateTable(
                name: "EvaluationRecords",
                columns: table => new
                {
                    Id = table.Column<int>(type: "INTEGER", nullable: false)
                        .Annotation("Sqlite:Autoincrement", true),
                    PersonnelId = table.Column<int>(type: "INTEGER", nullable: false),
                    Year = table.Column<int>(type: "INTEGER", nullable: false),
                    Rating = table.Column<string>(type: "TEXT", nullable: false),
                    DecisionNumber = table.Column<string>(type: "TEXT", nullable: true),
                    DecisionDate = table.Column<DateTime>(type: "TEXT", nullable: true),
                    DecisionAgency = table.Column<string>(type: "TEXT", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_EvaluationRecords", x => x.Id);
                    table.ForeignKey(
                        name: "FK_EvaluationRecords_Personnel_PersonnelId",
                        column: x => x.PersonnelId,
                        principalTable: "Personnel",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateTable(
                name: "PersonnelDegrees",
                columns: table => new
                {
                    Id = table.Column<int>(type: "INTEGER", nullable: false)
                        .Annotation("Sqlite:Autoincrement", true),
                    PersonnelId = table.Column<int>(type: "INTEGER", nullable: false),
                    DegreeType = table.Column<string>(type: "TEXT", nullable: false),
                    DegreeName = table.Column<string>(type: "TEXT", nullable: false),
                    Major = table.Column<string>(type: "TEXT", nullable: true),
                    Institution = table.Column<string>(type: "TEXT", nullable: true),
                    Classification = table.Column<string>(type: "TEXT", nullable: true),
                    GraduationYear = table.Column<string>(type: "TEXT", nullable: true),
                    IssueDate = table.Column<DateTime>(type: "TEXT", nullable: true),
                    TrainingForm = table.Column<string>(type: "TEXT", nullable: true),
                    IsPrimary = table.Column<bool>(type: "INTEGER", nullable: false),
                    Note = table.Column<string>(type: "TEXT", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_PersonnelDegrees", x => x.Id);
                    table.ForeignKey(
                        name: "FK_PersonnelDegrees_Personnel_PersonnelId",
                        column: x => x.PersonnelId,
                        principalTable: "Personnel",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateTable(
                name: "PlanningPositions",
                columns: table => new
                {
                    Id = table.Column<int>(type: "INTEGER", nullable: false)
                        .Annotation("Sqlite:Autoincrement", true),
                    Name = table.Column<string>(type: "TEXT", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_PlanningPositions", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "PlanningRecords",
                columns: table => new
                {
                    Id = table.Column<int>(type: "INTEGER", nullable: false)
                        .Annotation("Sqlite:Autoincrement", true),
                    PersonnelId = table.Column<int>(type: "INTEGER", nullable: false),
                    PlanningType = table.Column<string>(type: "TEXT", nullable: false),
                    Status = table.Column<string>(type: "TEXT", nullable: false),
                    CurrentPosition = table.Column<string>(type: "TEXT", nullable: true),
                    PlannedPosition = table.Column<string>(type: "TEXT", nullable: true),
                    PlannedTransitionPosition = table.Column<string>(type: "TEXT", nullable: true),
                    TrainingLevel = table.Column<string>(type: "TEXT", nullable: true),
                    PoliticalTheoryLevel = table.Column<string>(type: "TEXT", nullable: true),
                    PlanningTerm = table.Column<string>(type: "TEXT", nullable: true),
                    DecisionNumber = table.Column<string>(type: "TEXT", nullable: true),
                    DecisionDate = table.Column<DateTime>(type: "TEXT", nullable: true),
                    DecisionUnit = table.Column<string>(type: "TEXT", nullable: true),
                    Evaluation3Years = table.Column<string>(type: "TEXT", nullable: true),
                    Note = table.Column<string>(type: "TEXT", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_PlanningRecords", x => x.Id);
                    table.ForeignKey(
                        name: "FK_PlanningRecords_Personnel_PersonnelId",
                        column: x => x.PersonnelId,
                        principalTable: "Personnel",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateTable(
                name: "PlanningTerms",
                columns: table => new
                {
                    Id = table.Column<int>(type: "INTEGER", nullable: false)
                        .Annotation("Sqlite:Autoincrement", true),
                    TermName = table.Column<string>(type: "TEXT", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_PlanningTerms", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "RotationRecords",
                columns: table => new
                {
                    Id = table.Column<int>(type: "INTEGER", nullable: false)
                        .Annotation("Sqlite:Autoincrement", true),
                    PersonnelId = table.Column<int>(type: "INTEGER", nullable: false),
                    FromDepartment = table.Column<string>(type: "TEXT", nullable: true),
                    ToDepartment = table.Column<string>(type: "TEXT", nullable: true),
                    RotationType = table.Column<string>(type: "TEXT", nullable: false),
                    PlanType = table.Column<string>(type: "TEXT", nullable: false),
                    IsCompleted = table.Column<bool>(type: "INTEGER", nullable: false),
                    DecisionNumber = table.Column<string>(type: "TEXT", nullable: true),
                    DecisionDate = table.Column<DateTime>(type: "TEXT", nullable: true),
                    EffectiveDate = table.Column<DateTime>(type: "TEXT", nullable: true),
                    Note = table.Column<string>(type: "TEXT", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_RotationRecords", x => x.Id);
                    table.ForeignKey(
                        name: "FK_RotationRecords_Personnel_PersonnelId",
                        column: x => x.PersonnelId,
                        principalTable: "Personnel",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateTable(
                name: "TrainingClasses",
                columns: table => new
                {
                    Id = table.Column<int>(type: "INTEGER", nullable: false)
                        .Annotation("Sqlite:Autoincrement", true),
                    ClassName = table.Column<string>(type: "TEXT", nullable: false),
                    ParticipationDate = table.Column<DateTime>(type: "TEXT", nullable: true),
                    DecisionNumber = table.Column<string>(type: "TEXT", nullable: true),
                    DecisionDate = table.Column<DateTime>(type: "TEXT", nullable: true),
                    DecisionUnit = table.Column<string>(type: "TEXT", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_TrainingClasses", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "PersonnelTrainings",
                columns: table => new
                {
                    Id = table.Column<int>(type: "INTEGER", nullable: false)
                        .Annotation("Sqlite:Autoincrement", true),
                    PersonnelId = table.Column<int>(type: "INTEGER", nullable: false),
                    TrainingClassId = table.Column<int>(type: "INTEGER", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_PersonnelTrainings", x => x.Id);
                    table.ForeignKey(
                        name: "FK_PersonnelTrainings_Personnel_PersonnelId",
                        column: x => x.PersonnelId,
                        principalTable: "Personnel",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                    table.ForeignKey(
                        name: "FK_PersonnelTrainings_TrainingClasses_TrainingClassId",
                        column: x => x.TrainingClassId,
                        principalTable: "TrainingClasses",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_EvaluationRecords_PersonnelId",
                table: "EvaluationRecords",
                column: "PersonnelId");

            migrationBuilder.CreateIndex(
                name: "IX_PersonnelDegrees_PersonnelId",
                table: "PersonnelDegrees",
                column: "PersonnelId");

            migrationBuilder.CreateIndex(
                name: "IX_PersonnelTrainings_PersonnelId",
                table: "PersonnelTrainings",
                column: "PersonnelId");

            migrationBuilder.CreateIndex(
                name: "IX_PersonnelTrainings_TrainingClassId",
                table: "PersonnelTrainings",
                column: "TrainingClassId");

            migrationBuilder.CreateIndex(
                name: "IX_PlanningRecords_PersonnelId",
                table: "PlanningRecords",
                column: "PersonnelId");

            migrationBuilder.CreateIndex(
                name: "IX_RotationRecords_PersonnelId",
                table: "RotationRecords",
                column: "PersonnelId");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "EvaluationRecords");

            migrationBuilder.DropTable(
                name: "PersonnelDegrees");

            migrationBuilder.DropTable(
                name: "PersonnelTrainings");

            migrationBuilder.DropTable(
                name: "PlanningPositions");

            migrationBuilder.DropTable(
                name: "PlanningRecords");

            migrationBuilder.DropTable(
                name: "PlanningTerms");

            migrationBuilder.DropTable(
                name: "RotationRecords");

            migrationBuilder.DropTable(
                name: "TrainingClasses");

            migrationBuilder.DropColumn(
                name: "ExceedFramePercent",
                table: "SalaryRecords");

            migrationBuilder.DropColumn(
                name: "RankCode",
                table: "SalaryRecords");

            migrationBuilder.DropColumn(
                name: "SalaryStep",
                table: "SalaryRecords");

            migrationBuilder.DropColumn(
                name: "DepartmentName",
                table: "Positions");

            migrationBuilder.DropColumn(
                name: "GroupType",
                table: "Positions");

            migrationBuilder.DropColumn(
                name: "CurrentResidence",
                table: "Personnel");

            migrationBuilder.DropColumn(
                name: "Hometown",
                table: "Personnel");
        }
    }
}
