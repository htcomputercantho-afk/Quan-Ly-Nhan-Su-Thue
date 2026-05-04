using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace TaxPersonnelManagement.Migrations
{
    /// <inheritdoc />
    public partial class AddSalaryRecordColumns : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "EffectiveDate",
                table: "SalaryRecords");

            migrationBuilder.DropColumn(
                name: "NextIncreaseDate",
                table: "SalaryRecords");

            migrationBuilder.DropColumn(
                name: "SalaryLevel",
                table: "SalaryRecords");

            migrationBuilder.AlterColumn<DateTime>(
                name: "DecisionDate",
                table: "SalaryRecords",
                type: "TEXT",
                nullable: true,
                oldClrType: typeof(DateTime),
                oldType: "TEXT");

            migrationBuilder.AlterColumn<string>(
                name: "Coefficient",
                table: "SalaryRecords",
                type: "TEXT",
                nullable: true,
                oldClrType: typeof(decimal),
                oldType: "TEXT");

            migrationBuilder.AddColumn<string>(
                name: "DecisionNumber",
                table: "SalaryRecords",
                type: "TEXT",
                nullable: true);

            migrationBuilder.AddColumn<DateTime>(
                name: "EndDate",
                table: "SalaryRecords",
                type: "TEXT",
                nullable: true);

            migrationBuilder.AddColumn<double>(
                name: "Percentage",
                table: "SalaryRecords",
                type: "REAL",
                nullable: false,
                defaultValue: 0.0);

            migrationBuilder.AddColumn<DateTime>(
                name: "SalaryCalculationDate",
                table: "SalaryRecords",
                type: "TEXT",
                nullable: true);

            migrationBuilder.AddColumn<DateTime>(
                name: "StartDate",
                table: "SalaryRecords",
                type: "TEXT",
                nullable: true);

            migrationBuilder.AddColumn<DateTime>(
                name: "DisciplineDecisionDate",
                table: "Personnel",
                type: "TEXT",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "DisciplineDecisionNumber",
                table: "Personnel",
                type: "TEXT",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "DisciplineReason",
                table: "Personnel",
                type: "TEXT",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "DisciplineType",
                table: "Personnel",
                type: "TEXT",
                nullable: true);

            migrationBuilder.AlterColumn<DateTime>(
                name: "EndDate",
                table: "LeaveHistories",
                type: "TEXT",
                nullable: true,
                oldClrType: typeof(DateTime),
                oldType: "TEXT");

            migrationBuilder.AddColumn<int>(
                name: "LeaveYear",
                table: "LeaveHistories",
                type: "INTEGER",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "SystemNote",
                table: "LeaveHistories",
                type: "TEXT",
                nullable: true);

            migrationBuilder.CreateTable(
                name: "DisciplineTypes",
                columns: table => new
                {
                    Id = table.Column<int>(type: "INTEGER", nullable: false)
                        .Annotation("Sqlite:Autoincrement", true),
                    Name = table.Column<string>(type: "TEXT", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_DisciplineTypes", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "IncomeRecords",
                columns: table => new
                {
                    Id = table.Column<int>(type: "INTEGER", nullable: false)
                        .Annotation("Sqlite:Autoincrement", true),
                    PersonnelId = table.Column<int>(type: "INTEGER", nullable: false),
                    Year = table.Column<int>(type: "INTEGER", nullable: false),
                    Month = table.Column<int>(type: "INTEGER", nullable: false),
                    IncomeType = table.Column<string>(type: "TEXT", nullable: false),
                    Amount = table.Column<decimal>(type: "TEXT", nullable: false),
                    Note = table.Column<string>(type: "TEXT", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_IncomeRecords", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "PublicHolidays",
                columns: table => new
                {
                    Id = table.Column<int>(type: "INTEGER", nullable: false)
                        .Annotation("Sqlite:Autoincrement", true),
                    Name = table.Column<string>(type: "TEXT", nullable: false),
                    Date = table.Column<DateTime>(type: "TEXT", nullable: false),
                    Note = table.Column<string>(type: "TEXT", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_PublicHolidays", x => x.Id);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "DisciplineTypes");

            migrationBuilder.DropTable(
                name: "IncomeRecords");

            migrationBuilder.DropTable(
                name: "PublicHolidays");

            migrationBuilder.DropColumn(
                name: "DecisionNumber",
                table: "SalaryRecords");

            migrationBuilder.DropColumn(
                name: "EndDate",
                table: "SalaryRecords");

            migrationBuilder.DropColumn(
                name: "Percentage",
                table: "SalaryRecords");

            migrationBuilder.DropColumn(
                name: "SalaryCalculationDate",
                table: "SalaryRecords");

            migrationBuilder.DropColumn(
                name: "StartDate",
                table: "SalaryRecords");

            migrationBuilder.DropColumn(
                name: "DisciplineDecisionDate",
                table: "Personnel");

            migrationBuilder.DropColumn(
                name: "DisciplineDecisionNumber",
                table: "Personnel");

            migrationBuilder.DropColumn(
                name: "DisciplineReason",
                table: "Personnel");

            migrationBuilder.DropColumn(
                name: "DisciplineType",
                table: "Personnel");

            migrationBuilder.DropColumn(
                name: "LeaveYear",
                table: "LeaveHistories");

            migrationBuilder.DropColumn(
                name: "SystemNote",
                table: "LeaveHistories");

            migrationBuilder.AlterColumn<DateTime>(
                name: "DecisionDate",
                table: "SalaryRecords",
                type: "TEXT",
                nullable: false,
                defaultValue: new DateTime(1, 1, 1, 0, 0, 0, 0, DateTimeKind.Unspecified),
                oldClrType: typeof(DateTime),
                oldType: "TEXT",
                oldNullable: true);

            migrationBuilder.AlterColumn<decimal>(
                name: "Coefficient",
                table: "SalaryRecords",
                type: "TEXT",
                nullable: false,
                defaultValue: 0m,
                oldClrType: typeof(string),
                oldType: "TEXT",
                oldNullable: true);

            migrationBuilder.AddColumn<DateTime>(
                name: "EffectiveDate",
                table: "SalaryRecords",
                type: "TEXT",
                nullable: false,
                defaultValue: new DateTime(1, 1, 1, 0, 0, 0, 0, DateTimeKind.Unspecified));

            migrationBuilder.AddColumn<DateTime>(
                name: "NextIncreaseDate",
                table: "SalaryRecords",
                type: "TEXT",
                nullable: false,
                defaultValue: new DateTime(1, 1, 1, 0, 0, 0, 0, DateTimeKind.Unspecified));

            migrationBuilder.AddColumn<string>(
                name: "SalaryLevel",
                table: "SalaryRecords",
                type: "TEXT",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AlterColumn<DateTime>(
                name: "EndDate",
                table: "LeaveHistories",
                type: "TEXT",
                nullable: false,
                defaultValue: new DateTime(1, 1, 1, 0, 0, 0, 0, DateTimeKind.Unspecified),
                oldClrType: typeof(DateTime),
                oldType: "TEXT",
                oldNullable: true);
        }
    }
}
