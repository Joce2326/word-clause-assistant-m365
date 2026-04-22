export interface IClauseTemplate {
    id: string;
    category: string;
    title: string;
    content: string;
}

export const clauses: IClauseTemplate[] = [
    {
        id: "1",
        category: "Confidentiality",
        title: "Standard Confidentiality Clause",
        content:
            "The parties agree to keep all confidential information strictly confidential and not to disclose such information to any third party except as required by law."
    },
    {
        id: "2",
        category: "Liability",
        title: "Limitation of Liability Clause",
        content:
            "Neither party shall be liable for any indirect, incidental, special, or consequential damages arising out of or in connection with this Agreement."
    },
    {
        id: "3",
        category: "Termination",
        title: "Termination for Convenience",
        content:
            "Either party may terminate this Agreement upon thirty (30) days written notice to the other party."
    }
];